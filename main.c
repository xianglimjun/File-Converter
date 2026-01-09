#define _CRT_SECURE_NO_WARNINGS
#include <windows.h>
#include <commctrl.h>
#include <commdlg.h>
#include <stdio.h>
#include <string.h>
#include <process.h>

#pragma comment(lib, "comctl32.lib")

// Image libraries
#define STB_IMAGE_IMPLEMENTATION
#include "stb_image.h"

#define STB_IMAGE_WRITE_IMPLEMENTATION
#include "stb_image_write.h"

// WebP support
#include <webp/encode.h>
#include <webp/decode.h>

// Animated GIF support
#define MSF_GIF_IMPL
#include "msf_gif.h"

// Custom PDF writer
#include "pdf_writer.h"

#ifndef SS_ELLIPSISPATH
#define SS_ELLIPSISPATH 0x00004000L
#endif

// ============================================================================
// Embedded VBScript for Office conversion
// ============================================================================
static const char* OFFICE_VBS_SCRIPT = 
"Option Explicit\n"
"Dim fso, inputFile, outputFile, ext\n"
"Dim wordApp, doc, pptApp, pres\n"
"If WScript.Arguments.Count < 2 Then\n"
"    WScript.Quit 1\n"
"End If\n"
"inputFile = WScript.Arguments(0)\n"
"outputFile = WScript.Arguments(1)\n"
"Set fso = CreateObject(\"Scripting.FileSystemObject\")\n"
"inputFile = fso.GetAbsolutePathName(inputFile)\n"
"outputFile = fso.GetAbsolutePathName(outputFile)\n"
"ext = LCase(fso.GetExtensionName(inputFile))\n"
"On Error Resume Next\n"
"If ext = \"docx\" Or ext = \"doc\" Then\n"
"    Set wordApp = CreateObject(\"Word.Application\")\n"
"    If Err.Number <> 0 Then WScript.Quit 1\n"
"    wordApp.Visible = False\n"
"    wordApp.DisplayAlerts = 0\n"
"    Set doc = wordApp.Documents.Open(inputFile)\n"
"    If Err.Number <> 0 Then wordApp.Quit: WScript.Quit 1\n"
"    doc.SaveAs outputFile, 17\n"
"    If Err.Number <> 0 Then doc.Close 0: wordApp.Quit: WScript.Quit 1\n"
"    doc.Close 0\n"
"    wordApp.Quit\n"
"ElseIf ext = \"pptx\" Or ext = \"ppt\" Then\n"
"    Set pptApp = CreateObject(\"PowerPoint.Application\")\n"
"    If Err.Number <> 0 Then WScript.Quit 1\n"
"    pptApp.DisplayAlerts = 1\n"
"    Set pres = pptApp.Presentations.Open(inputFile, -1, 0, 0)\n"
"    If Err.Number <> 0 Then pptApp.Quit: WScript.Quit 1\n"
"    pres.SaveAs outputFile, 32\n"
"    If Err.Number <> 0 Then pres.Close: pptApp.Quit: WScript.Quit 1\n"
"    pres.Close\n"
"    pptApp.Quit\n"
"Else\n"
"    WScript.Quit 1\n"
"End If\n"
"If Err.Number <> 0 Then WScript.Quit 1\n"
"WScript.Quit 0\n";

int GetTempVbsPath(char* outPath, size_t maxLen) {
    char tempDir[MAX_PATH];
    if (GetTempPath(MAX_PATH, tempDir) == 0) return 0;
    snprintf(outPath, maxLen, "%s_office_conv.vbs", tempDir);
    FILE* f = fopen(outPath, "w");
    if (!f) return 0;
    fputs(OFFICE_VBS_SCRIPT, f);
    fclose(f);
    return 1;
}

// ============================================================================
// GUI Constants
// ============================================================================
#define ID_BTN_SELECT    101
#define ID_LBL_PATH      102
#define ID_GRP_OPTIONS   103
#define ID_CMB_FORMAT    104
#define ID_CHK_COPY      105
#define ID_BTN_CONVERT   106
#define ID_PROGRESS      107
#define ID_LBL_STATUS    108

// Custom messages for thread communication
#define WM_CONVERT_PROGRESS (WM_USER + 1)
#define WM_CONVERT_DONE     (WM_USER + 2)

// Format indices
#define FMT_PNG  0
#define FMT_JPG  1
#define FMT_WEBP 2
#define FMT_PDF  3
#define FMT_GIF  4  // Only for multi-image

// ============================================================================
// Global State
// ============================================================================
HWND hMainWnd, hLabelPath, hComboFormat, hCheckCopy, hProgress, hBtnConvert, hBtnSelect, hLabelStatus;
HFONT hFont;
int isConverting = 0;  // Flag to prevent double-click

#define MAX_FILES 100
char selectedFiles[MAX_FILES][MAX_PATH];
int fileCount = 0;
int allOffice = 0;  // True if all selected files are Office docs
int g_formatIndex = 0;
int g_saveCopy = 0;

// ============================================================================
// Helpers
// ============================================================================
void SetControlFont(HWND hwnd) {
    SendMessage(hwnd, WM_SETFONT, (WPARAM)hFont, TRUE);
}

int HasExtension(const char* filename, const char* ext) {
    const char* dot = strrchr(filename, '.');
    if (!dot) return 0;
    return _stricmp(dot, ext) == 0;
}

int IsOfficeFile(const char* filename) {
    return HasExtension(filename, ".docx") || HasExtension(filename, ".doc") ||
           HasExtension(filename, ".pptx") || HasExtension(filename, ".ppt");
}

int IsImageFile(const char* filename) {
    return HasExtension(filename, ".jpg") || HasExtension(filename, ".jpeg") ||
           HasExtension(filename, ".png") || HasExtension(filename, ".gif") ||
           HasExtension(filename, ".webp");
}

// ============================================================================
// WebP Decode Helper
// ============================================================================
unsigned char* LoadWebP(const char* filename, int* width, int* height, int* channels) {
    FILE* f = fopen(filename, "rb");
    if (!f) return NULL;
    
    fseek(f, 0, SEEK_END);
    size_t size = ftell(f);
    fseek(f, 0, SEEK_SET);
    
    uint8_t* webpData = (uint8_t*)malloc(size);
    if (!webpData) { fclose(f); return NULL; }
    
    fread(webpData, 1, size, f);
    fclose(f);
    
    unsigned char* rgba = WebPDecodeRGBA(webpData, size, width, height);
    free(webpData);
    
    if (rgba) {
        *channels = 4;
        return rgba;
    }
    return NULL;
}

// ============================================================================
// WebP Write Helper
// ============================================================================
int WriteWebP(const char* filename, int width, int height, int channels, unsigned char* data) {
    uint8_t* output = NULL;
    size_t size = 0;
    
    // WebP needs RGB or RGBA - if we have grayscale, expand it
    if (channels == 4) {
        size = WebPEncodeRGBA(data, width, height, width * 4, 85, &output);
    } else if (channels == 3) {
        size = WebPEncodeRGB(data, width, height, width * 3, 85, &output);
    } else if (channels == 1 || channels == 2) {
        // Convert grayscale to RGB/RGBA
        int newChannels = (channels == 2) ? 4 : 3;
        unsigned char* rgb = (unsigned char*)malloc(width * height * newChannels);
        if (!rgb) return 0;
        
        for (int i = 0; i < width * height; i++) {
            rgb[i * newChannels + 0] = data[i * channels];
            rgb[i * newChannels + 1] = data[i * channels];
            rgb[i * newChannels + 2] = data[i * channels];
            if (newChannels == 4) {
                rgb[i * newChannels + 3] = (channels == 2) ? data[i * channels + 1] : 255;
            }
        }
        
        if (newChannels == 4) {
            size = WebPEncodeRGBA(rgb, width, height, width * 4, 85, &output);
        } else {
            size = WebPEncodeRGB(rgb, width, height, width * 3, 85, &output);
        }
        free(rgb);
    } else {
        return 0;
    }
    
    if (size == 0 || output == NULL) return 0;
    
    FILE* f = fopen(filename, "wb");
    if (!f) {
        WebPFree(output);
        return 0;
    }
    fwrite(output, 1, size, f);
    fclose(f);
    WebPFree(output);
    return 1;
}

// ============================================================================
// Office Document Conversion
// ============================================================================
int ConvertOfficeDoc(const char* srcPath, const char* destPath) {
    char vbsPath[MAX_PATH];
    if (!GetTempVbsPath(vbsPath, sizeof(vbsPath))) return 0;

    char cmd[MAX_PATH * 4];
    snprintf(cmd, sizeof(cmd), "cscript //Nologo \"%s\" \"%s\" \"%s\"", vbsPath, srcPath, destPath);
    
    STARTUPINFO si = {0};
    PROCESS_INFORMATION pi = {0};
    si.cb = sizeof(si);
    si.dwFlags = STARTF_USESHOWWINDOW;
    si.wShowWindow = SW_HIDE;
    
    int ret = 1;
    if (CreateProcess(NULL, cmd, NULL, NULL, FALSE, CREATE_NO_WINDOW, NULL, NULL, &si, &pi)) {
        WaitForSingleObject(pi.hProcess, INFINITE);
        DWORD exitCode;
        GetExitCodeProcess(pi.hProcess, &exitCode);
        ret = (int)exitCode;
        CloseHandle(pi.hProcess);
        CloseHandle(pi.hThread);
    }
    
    DeleteFile(vbsPath);
    return (ret == 0);
}

// ============================================================================
// Image Conversion
// ============================================================================
int ConvertImage(const char* srcPath, const char* destPath, int formatIndex) {
    int width, height, channels;
    unsigned char* img = NULL;
    int isWebP = HasExtension(srcPath, ".webp");
    
    if (isWebP) {
        img = LoadWebP(srcPath, &width, &height, &channels);
    } else {
        img = stbi_load(srcPath, &width, &height, &channels, 0);
    }
    
    if (!img) return 0;
    
    int result = 0;
    switch (formatIndex) {
        case FMT_PNG:
            result = stbi_write_png(destPath, width, height, channels, img, width * channels);
            break;
        case FMT_JPG:
            result = stbi_write_jpg(destPath, width, height, channels, img, 90);
            break;
        case FMT_WEBP:
            result = WriteWebP(destPath, width, height, channels, img);
            break;
        case FMT_PDF:
            result = create_pdf_from_image(destPath, width, height, channels, img);
            break;
    }
    
    if (isWebP) {
        WebPFree(img); // LoadWebP returns result of WebPDecodeRGBA which needs WebPFree
    } else {
        stbi_image_free(img);
    }
    return result;
}

// ============================================================================
// Animated GIF Creation (Multi-image)
// ============================================================================
int CreateAnimatedGif(const char* destPath, int delayMs) {
    if (fileCount < 2) return 0;
    
    // Load first image to get dimensions
    int width, height, channels;
    unsigned char* firstImg = stbi_load(selectedFiles[0], &width, &height, &channels, 4);
    if (!firstImg) return 0;
    
    MsfGifState gifState = {0};
    msf_gif_begin(&gifState, width, height);
    msf_gif_frame(&gifState, firstImg, delayMs / 10, 16, width * 4);
    stbi_image_free(firstImg);
    
    // Add remaining frames
    for (int i = 1; i < fileCount; i++) {
        int w, h, c;
        unsigned char* img = stbi_load(selectedFiles[i], &w, &h, &c, 4);
        if (img) {
            // Resize if needed (simple skip for now - use same size)
            if (w == width && h == height) {
                msf_gif_frame(&gifState, img, delayMs / 10, 16, width * 4);
            }
            stbi_image_free(img);
        }
    }
    
    MsfGifResult result = msf_gif_end(&gifState);
    if (result.data) {
        FILE* f = fopen(destPath, "wb");
        if (f) {
            fwrite(result.data, 1, result.dataSize, f);
            fclose(f);
            msf_gif_free(result);
            return 1;
        }
        msf_gif_free(result);
    }
    return 0;
}

// ============================================================================
// Worker Thread for Conversion
// ============================================================================
typedef struct {
    int success;
    int failed;
} ConvertResult;

unsigned __stdcall ConvertWorkerThread(void* param) {
    int success = 0, failed = 0;
    
    for (int i = 0; i < fileCount; i++) {
        char destPath[MAX_PATH];
        strcpy(destPath, selectedFiles[i]);
        
        char* dot = strrchr(destPath, '.');
        if (dot) *dot = '\0';
        if (g_saveCopy) strcat(destPath, "_copy");
        
        int result = 0;
        if (IsOfficeFile(selectedFiles[i])) {
            strcat(destPath, ".pdf");
            result = ConvertOfficeDoc(selectedFiles[i], destPath);
        } else {
            const char* extensions[] = {".png", ".jpg", ".webp", ".pdf", ".gif"};
            strcat(destPath, extensions[g_formatIndex]);
            result = ConvertImage(selectedFiles[i], destPath, g_formatIndex);
        }
        
        if (result) success++;
        else failed++;
        
        // Update progress bar
        PostMessage(hMainWnd, WM_CONVERT_PROGRESS, i + 1, fileCount);
    }
    
    // Notify completion
    PostMessage(hMainWnd, WM_CONVERT_DONE, success, failed);
    return 0;
}

// ============================================================================
// Main Conversion Logic
// ============================================================================
void ConvertFiles(HWND hwnd) {
    if (fileCount == 0) {
        MessageBox(hwnd, "Please select file(s) first.", "Error", MB_OK | MB_ICONERROR);
        return;
    }
    
    if (isConverting) return;  // Prevent double-click
    
    g_formatIndex = SendMessage(hComboFormat, CB_GETCURSEL, 0, 0);
    g_saveCopy = (SendMessage(hCheckCopy, BM_GETCHECK, 0, 0) == BST_CHECKED);
    
    // Handle animated GIF (special case: multiple images -> single GIF)
    if (g_formatIndex == FMT_GIF && fileCount > 1) {
        char destPath[MAX_PATH];
        strcpy(destPath, selectedFiles[0]);
        char* dot = strrchr(destPath, '.');
        if (dot) *dot = '\0';
        if (g_saveCopy) strcat(destPath, "_copy");
        strcat(destPath, ".gif");
        
        if (CreateAnimatedGif(destPath, 500)) {
            MessageBox(hwnd, "Animated GIF created successfully!", "Success", MB_OK | MB_ICONINFORMATION);
        } else {
            MessageBox(hwnd, "Failed to create animated GIF.", "Error", MB_OK | MB_ICONERROR);
        }
        return;
    }
    
    // Disable controls during conversion
    isConverting = 1;
    EnableWindow(hBtnConvert, FALSE);
    EnableWindow(hBtnSelect, FALSE);
    EnableWindow(hComboFormat, FALSE);
    
    // Setup progress bar and status
    SendMessage(hProgress, PBM_SETRANGE, 0, MAKELPARAM(0, fileCount));
    SendMessage(hProgress, PBM_SETPOS, 0, 0);
    ShowWindow(hProgress, SW_SHOW);
    SetWindowText(hLabelStatus, "Starting conversion...");
    ShowWindow(hLabelStatus, SW_SHOW);
    
    // Start worker thread
    _beginthreadex(NULL, 0, ConvertWorkerThread, NULL, 0, NULL);
}

// ============================================================================
// File Selection (Multi-file support)
// ============================================================================
void SelectFiles(HWND hwnd) {
    static char szFile[MAX_FILES * MAX_PATH] = "";
    
    OPENFILENAME ofn = {0};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = hwnd;
    ofn.lpstrFile = szFile;
    ofn.nMaxFile = sizeof(szFile);
    ofn.lpstrFilter = "All Supported\0*.jpg;*.jpeg;*.png;*.gif;*.webp;*.docx;*.pptx;*.doc;*.ppt\0"
                      "Images\0*.jpg;*.jpeg;*.png;*.gif;*.webp\0"
                      "Office\0*.docx;*.pptx;*.doc;*.ppt\0"
                      "All Files\0*.*\0";
    ofn.nFilterIndex = 1;
    ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST | OFN_ALLOWMULTISELECT | OFN_EXPLORER;
    
    szFile[0] = '\0';
    
    if (GetOpenFileName(&ofn)) {
        fileCount = 0;
        allOffice = 1;
        
        // Parse multi-select result
        char* p = szFile;
        char dir[MAX_PATH] = "";
        
        if (ofn.nFileOffset > 0 && szFile[ofn.nFileOffset - 1] == '\0') {
            // Multiple files: first part is directory
            strcpy(dir, p);
            p += strlen(p) + 1;
            
            while (*p && fileCount < MAX_FILES) {
                snprintf(selectedFiles[fileCount], MAX_PATH, "%s\\%s", dir, p);
                if (!IsOfficeFile(selectedFiles[fileCount])) allOffice = 0;
                fileCount++;
                p += strlen(p) + 1;
            }
        } else {
            // Single file
            strcpy(selectedFiles[0], szFile);
            allOffice = IsOfficeFile(selectedFiles[0]);
            fileCount = 1;
        }
        
        // Update UI
        char label[128];
        if (fileCount == 1) {
            SetWindowText(hLabelPath, selectedFiles[0]);
        } else {
            snprintf(label, sizeof(label), "%d files selected", fileCount);
            SetWindowText(hLabelPath, label);
        }
        
        // Smart UI: Update format options
        SendMessage(hComboFormat, CB_RESETCONTENT, 0, 0);
        
        if (allOffice) {
            // Office docs: Only PDF allowed
            SendMessage(hComboFormat, CB_ADDSTRING, 0, (LPARAM)"PDF (.pdf)");
            SendMessage(hComboFormat, CB_SETCURSEL, 0, 0);
            EnableWindow(hComboFormat, FALSE);
        } else {
            // Images: Show all options
            SendMessage(hComboFormat, CB_ADDSTRING, 0, (LPARAM)"PNG (.png)");
            SendMessage(hComboFormat, CB_ADDSTRING, 0, (LPARAM)"JPG (.jpg)");
            SendMessage(hComboFormat, CB_ADDSTRING, 0, (LPARAM)"WebP (.webp)");
            SendMessage(hComboFormat, CB_ADDSTRING, 0, (LPARAM)"PDF (.pdf)");
            
            // GIF only for multiple images
            if (fileCount > 1) {
                SendMessage(hComboFormat, CB_ADDSTRING, 0, (LPARAM)"Animated GIF (.gif)");
            }
            
            SendMessage(hComboFormat, CB_SETCURSEL, 0, 0);
            EnableWindow(hComboFormat, TRUE);
        }
    }
}

// ============================================================================
// Window Procedure
// ============================================================================
LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
        case WM_CREATE: {
            hFont = CreateFont(16, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, ANSI_CHARSET, 
                               OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, 
                               DEFAULT_PITCH | FF_SWISS, "Segoe UI");

            HWND hBtnSelect = CreateWindow("BUTTON", "Select File(s)", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON, 
                         20, 20, 340, 30, hwnd, (HMENU)ID_BTN_SELECT, NULL, NULL);
            SetControlFont(hBtnSelect);

            hLabelPath = CreateWindow("STATIC", "No file selected", WS_VISIBLE | WS_CHILD | SS_CENTER | SS_ELLIPSISPATH, 
                                      20, 60, 340, 20, hwnd, (HMENU)ID_LBL_PATH, NULL, NULL);
            SetControlFont(hLabelPath);

            HWND hGrp = CreateWindow("BUTTON", "Options", WS_VISIBLE | WS_CHILD | BS_GROUPBOX, 
                         20, 90, 340, 100, hwnd, (HMENU)ID_GRP_OPTIONS, NULL, NULL);
            SetControlFont(hGrp);

            CreateWindow("STATIC", "Target Format:", WS_VISIBLE | WS_CHILD, 
                         40, 115, 100, 20, hwnd, NULL, NULL, NULL);

            hComboFormat = CreateWindow("COMBOBOX", "", WS_VISIBLE | WS_CHILD | CBS_DROPDOWNLIST, 
                                        140, 112, 200, 200, hwnd, (HMENU)ID_CMB_FORMAT, NULL, NULL);
            SetControlFont(hComboFormat);
            SendMessage(hComboFormat, CB_ADDSTRING, 0, (LPARAM)"PNG (.png)");
            SendMessage(hComboFormat, CB_ADDSTRING, 0, (LPARAM)"JPG (.jpg)");
            SendMessage(hComboFormat, CB_ADDSTRING, 0, (LPARAM)"WebP (.webp)");
            SendMessage(hComboFormat, CB_ADDSTRING, 0, (LPARAM)"PDF (.pdf)");
            SendMessage(hComboFormat, CB_SETCURSEL, 0, 0);

            hCheckCopy = CreateWindow("BUTTON", "Save as separate copy (_copy)", WS_VISIBLE | WS_CHILD | BS_AUTOCHECKBOX, 
                                      40, 150, 300, 20, hwnd, (HMENU)ID_CHK_COPY, NULL, NULL);
            SetControlFont(hCheckCopy);

            // Progress bar (initially hidden)
            hProgress = CreateWindowEx(0, PROGRESS_CLASS, "", WS_CHILD | PBS_SMOOTH,
                         20, 195, 340, 15, hwnd, (HMENU)ID_PROGRESS, NULL, NULL);

            // Status Label (initially hidden)
            hLabelStatus = CreateWindow("STATIC", "", WS_CHILD | SS_CENTER, 
                         20, 215, 340, 20, hwnd, (HMENU)ID_LBL_STATUS, NULL, NULL);
            SetControlFont(hLabelStatus);

            hBtnConvert = CreateWindow("BUTTON", "CONVERT", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON | BS_DEFPUSHBUTTON, 
                         20, 240, 340, 40, hwnd, (HMENU)ID_BTN_CONVERT, NULL, NULL);
            HFONT hFontBold = CreateFont(18, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE, ANSI_CHARSET, 
                               OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, 
                               DEFAULT_PITCH | FF_SWISS, "Segoe UI");
            SendMessage(hBtnConvert, WM_SETFONT, (WPARAM)hFontBold, TRUE);

            // Store button handle for later
            hBtnSelect = GetDlgItem(hwnd, ID_BTN_SELECT);
            hMainWnd = hwnd;

            break;
        }
        case WM_COMMAND: {
            if (LOWORD(wParam) == ID_BTN_SELECT) {
                SelectFiles(hwnd);
            } else if (LOWORD(wParam) == ID_BTN_CONVERT) {
                ConvertFiles(hwnd);
            }
            break;
        }
        case WM_CONVERT_PROGRESS: {
            // wParam = current, lParam = total
            int current = (int)wParam;
            int total = (int)lParam;
            SendMessage(hProgress, PBM_SETPOS, (WPARAM)current, 0);
            
            // Update status text
            char status[64];
            snprintf(status, sizeof(status), "Converting %d of %d files...", current, total);
            SetWindowText(hLabelStatus, status);
            break;
        }
        case WM_CONVERT_DONE: {
            // wParam = success count, lParam = failed count
            int success = (int)wParam;
            int failed = (int)lParam;
            
            // Hide progress and status, re-enable controls
            ShowWindow(hProgress, SW_HIDE);
            ShowWindow(hLabelStatus, SW_HIDE);
            EnableWindow(hBtnConvert, TRUE);
            EnableWindow(hBtnSelect, TRUE);
            EnableWindow(hComboFormat, !allOffice);
            isConverting = 0;
            
            char msg[256];
            if (failed == 0) {
                snprintf(msg, sizeof(msg), "Successfully converted %d file(s)!", success);
                MessageBox(hwnd, msg, "Success", MB_OK | MB_ICONINFORMATION);
            } else {
                snprintf(msg, sizeof(msg), "Converted %d file(s), %d failed.", success, failed);
                MessageBox(hwnd, msg, "Partial Success", MB_OK | MB_ICONWARNING);
            }
            break;
        }
        case WM_DESTROY:
            DeleteObject(hFont);
            PostQuitMessage(0);
            break;
        default:
            return DefWindowProc(hwnd, msg, wParam, lParam);
    }
    return 0;
}

// ============================================================================
// Entry Point
// ============================================================================
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    // Initialize common controls for progress bar
    INITCOMMONCONTROLSEX icex;
    icex.dwSize = sizeof(icex);
    icex.dwICC = ICC_PROGRESS_CLASS;
    InitCommonControlsEx(&icex);
    
    WNDCLASSEX wc = {0};
    wc.cbSize = sizeof(WNDCLASSEX);
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.hIcon = LoadIcon(NULL, IDI_APPLICATION);
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)GetStockObject(WHITE_BRUSH); // Modern White background
    wc.lpszClassName = "UniversalConverterClass";
    wc.hIconSm = LoadIcon(NULL, IDI_APPLICATION);

    if (!RegisterClassEx(&wc)) {
        MessageBox(NULL, "Window Registration Failed!", "Error", MB_ICONEXCLAMATION | MB_OK);
        return 0;
    }

    HWND hwnd = CreateWindowEx(
        0, "UniversalConverterClass", "Universal File Converter",
        WS_OVERLAPPEDWINDOW & ~WS_MAXIMIZEBOX & ~WS_THICKFRAME,
        CW_USEDEFAULT, CW_USEDEFAULT, 400, 330,
        NULL, NULL, hInstance, NULL);

    if (hwnd == NULL) {
        MessageBox(NULL, "Window Creation Failed!", "Error", MB_ICONEXCLAMATION | MB_OK);
        return 0;
    }

    ShowWindow(hwnd, nCmdShow);
    UpdateWindow(hwnd);

    MSG Msg;
    while (GetMessage(&Msg, NULL, 0, 0) > 0) {
        TranslateMessage(&Msg);
        DispatchMessage(&Msg);
    }
    return Msg.wParam;
}
