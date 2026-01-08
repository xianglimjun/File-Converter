#define _CRT_SECURE_NO_WARNINGS
#include <windows.h>
#include <commdlg.h>
#include <stdio.h>
#include <string.h>

#define STB_IMAGE_IMPLEMENTATION
#include "stb_image.h"

#define STB_IMAGE_WRITE_IMPLEMENTATION
#include "stb_image_write.h"

#include "pdf_writer.h"

#ifndef SS_ELLIPSISPATH
#define SS_ELLIPSISPATH 0x00004000L
#endif

// Embedded VBScript for Office conversion
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

// Write embedded VBS to temp file and return path
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

// GUI Constants
#define ID_BTN_SELECT    101
#define ID_LBL_PATH      102
#define ID_GRP_OPTIONS   103
#define ID_CMB_FORMAT    104
#define ID_CHK_COPY      105
#define ID_BTN_CONVERT   106

// Global State
HWND hLabelPath, hComboFormat, hCheckCopy;
char sourcePath[MAX_PATH] = "";
HFONT hFont;

// Set nice font for a control
void SetFont(HWND hwnd) {
    SendMessage(hwnd, WM_SETFONT, (WPARAM)hFont, TRUE);
}

// Function prototypes
void ConvertFile(HWND hwnd);
void SelectFile(HWND hwnd);

// Helper to check extension
int HasExtension(const char* filename, const char* ext) {
    const char* dot = strrchr(filename, '.');
    if (!dot) return 0;
    return _stricmp(dot, ext) == 0;
}

void ConvertFile(HWND hwnd) {
    if (strlen(sourcePath) == 0) {
        MessageBox(hwnd, "Please select a file first.", "Error", MB_OK | MB_ICONERROR);
        return;
    }

    // Determine Output Format
    int formatIndex = SendMessage(hComboFormat, CB_GETCURSEL, 0, 0);
    // 0: PNG, 1: JPG, 2: BMP, 3: TGA, 4: PDF

    // Determine Output Filename
    char destPath[MAX_PATH];
    strcpy(destPath, sourcePath);
    
    // Strip extension
    char *dot = strrchr(destPath, '.');
    if (dot) *dot = '\0';

    // Append _copy if checked
    if (SendMessage(hCheckCopy, BM_GETCHECK, 0, 0) == BST_CHECKED) {
        strcat(destPath, "_copy");
    }

    // Append new extension
    switch (formatIndex) {
        case 0: strcat(destPath, ".png"); break;
        case 1: strcat(destPath, ".jpg"); break;
        case 2: strcat(destPath, ".bmp"); break;
        case 3: strcat(destPath, ".tga"); break;
        case 4: strcat(destPath, ".pdf"); break;
    }

    // Check if Input is Office Doc
    if (HasExtension(sourcePath, ".docx") || HasExtension(sourcePath, ".doc") || 
        HasExtension(sourcePath, ".pptx") || HasExtension(sourcePath, ".ppt")) {
        
        // Force PDF check (though UI should have handled it)
        if (formatIndex != 4) {
             MessageBox(hwnd, "Documents can only be converted to PDF.", "Warning", MB_OK | MB_ICONWARNING);
             strcpy(destPath + strlen(destPath) - 4, ".pdf");
        }

        // Write embedded VBS to temp and get path
        char vbsPath[MAX_PATH];
        if (!GetTempVbsPath(vbsPath, sizeof(vbsPath))) {
            MessageBox(hwnd, "Failed to create temp script.", "Error", MB_OK | MB_ICONERROR);
            return;
        }

        char cmd[MAX_PATH * 4];
        snprintf(cmd, sizeof(cmd), "cscript //Nologo \"%s\" \"%s\" \"%s\"", vbsPath, sourcePath, destPath);
        
        // Run hidden using CreateProcess
        STARTUPINFO si;
        PROCESS_INFORMATION pi;
        ZeroMemory(&si, sizeof(si));
        si.cb = sizeof(si);
        si.dwFlags = STARTF_USESHOWWINDOW;
        si.wShowWindow = SW_HIDE;
        ZeroMemory(&pi, sizeof(pi));
        
        int ret = 1; // Assume failure
        if (CreateProcess(NULL, cmd, NULL, NULL, FALSE, CREATE_NO_WINDOW, NULL, NULL, &si, &pi)) {
            WaitForSingleObject(pi.hProcess, INFINITE);
            DWORD exitCode;
            GetExitCodeProcess(pi.hProcess, &exitCode);
            ret = (int)exitCode;
            CloseHandle(pi.hProcess);
            CloseHandle(pi.hThread);
        }
        
        if (ret == 0) {
             MessageBox(hwnd, "Document converted successfully!", "Success", MB_OK | MB_ICONINFORMATION);
        } else {
             MessageBox(hwnd, "Document conversion failed.\nMake sure Microsoft 365 (Word/PowerPoint) is installed.", "Error", MB_OK | MB_ICONERROR);
        }
        
        // Clean up temp VBS file
        DeleteFile(vbsPath);
        return;
    }

    // Image Logic (Existing)
    int width, height, channels;
    unsigned char *img = stbi_load(sourcePath, &width, &height, &channels, 0);
    if (img == NULL) {
        MessageBox(hwnd, "Failed to load file. Passable causes:\n1. Not a supported image.\n2. Not a valid path.", "Error", MB_OK | MB_ICONERROR);
        return;
    }

    // Convert
    int result = 0;
    switch (formatIndex) {
        case 0: result = stbi_write_png(destPath, width, height, channels, img, width * channels); break;
        case 1: result = stbi_write_jpg(destPath, width, height, channels, img, 90); break;
        case 2: result = stbi_write_bmp(destPath, width, height, channels, img); break;
        case 3: result = stbi_write_tga(destPath, width, height, channels, img); break;
        case 4: result = create_pdf_from_image(destPath, width, height, channels, img); break;
    }

    stbi_image_free(img);

    if (result) {
        char msg[512];
        snprintf(msg, sizeof(msg), "Successfully converted to:\n%s", destPath);
        MessageBox(hwnd, "Success", msg, MB_OK | MB_ICONINFORMATION);
    } else {
        MessageBox(hwnd, "Conversion failed.", "Error", MB_OK | MB_ICONERROR);
    }
}

void SelectFile(HWND hwnd) {
    OPENFILENAME ofn;
    char szFile[MAX_PATH] = "";

    ZeroMemory(&ofn, sizeof(ofn));
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = hwnd;
    ofn.lpstrFile = szFile;
    ofn.nMaxFile = sizeof(szFile);
    ofn.lpstrFilter = "All Supported Files\0*.jpg;*.png;*.bmp;*.docx;*.pptx;*.doc;*.ppt\0Image Files\0*.jpg;*.png;*.bmp;*.tga;*.gif\0Office Files\0*.docx;*.pptx\0All Files\0*.*\0";
    ofn.nFilterIndex = 1;
    ofn.lpstrFileTitle = NULL;
    ofn.nMaxFileTitle = 0;
    ofn.lpstrInitialDir = NULL;
    ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;

    if (GetOpenFileName(&ofn) == TRUE) {
        strcpy(sourcePath, szFile);
        SetWindowText(hLabelPath, sourcePath);

        // Auto-Detect Document -> Set to PDF
        if (HasExtension(sourcePath, ".docx") || HasExtension(sourcePath, ".doc") || 
            HasExtension(sourcePath, ".pptx") || HasExtension(sourcePath, ".ppt")) {
            SendMessage(hComboFormat, CB_SETCURSEL, 4, 0); // Select PDF
        }
    }
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
        case WM_CREATE: {
            // Create Font
            hFont = CreateFont(16, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, ANSI_CHARSET, 
                               OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, 
                               DEFAULT_PITCH | FF_SWISS, "Segoe UI");

            // 1. Select File Button
            HWND hBtnSelect = CreateWindow("BUTTON", "Select File", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON, 
                         20, 20, 340, 30, hwnd, (HMENU)ID_BTN_SELECT, NULL, NULL);
            SetFont(hBtnSelect);

            // 2. File Path Label
            hLabelPath = CreateWindow("STATIC", "No file selected", WS_VISIBLE | WS_CHILD | SS_CENTER | SS_ELLIPSISPATH, 
                                      20, 60, 340, 20, hwnd, (HMENU)ID_LBL_PATH, NULL, NULL);
            SetFont(hLabelPath);

            // 3. Group Box
            HWND hGrp = CreateWindow("BUTTON", "Options", WS_VISIBLE | WS_CHILD | BS_GROUPBOX, 
                         20, 90, 340, 100, hwnd, (HMENU)ID_GRP_OPTIONS, NULL, NULL);
            SetFont(hGrp);

            // 4. Format Combobox
            CreateWindow("STATIC", "Target Format:", WS_VISIBLE | WS_CHILD, 
                         40, 115, 100, 20, hwnd, NULL, NULL, NULL); // Label

            hComboFormat = CreateWindow("COMBOBOX", "", WS_VISIBLE | WS_CHILD | CBS_DROPDOWNLIST, 
                                        140, 112, 200, 200, hwnd, (HMENU)ID_CMB_FORMAT, NULL, NULL);
            SetFont(hComboFormat);
            SendMessage(hComboFormat, CB_ADDSTRING, 0, (LPARAM)"PNG (.png)");
            SendMessage(hComboFormat, CB_ADDSTRING, 0, (LPARAM)"JPG (.jpg)");
            SendMessage(hComboFormat, CB_ADDSTRING, 0, (LPARAM)"BMP (.bmp)");
            SendMessage(hComboFormat, CB_ADDSTRING, 0, (LPARAM)"TGA (.tga)");
            SendMessage(hComboFormat, CB_ADDSTRING, 0, (LPARAM)"PDF (.pdf)");
            SendMessage(hComboFormat, CB_SETCURSEL, 0, 0); // Default to PNG

            // 5. Save as Copy Checkbox
            hCheckCopy = CreateWindow("BUTTON", "Save as separate copy (_copy)", WS_VISIBLE | WS_CHILD | BS_AUTOCHECKBOX, 
                                      40, 150, 300, 20, hwnd, (HMENU)ID_CHK_COPY, NULL, NULL);
            SetFont(hCheckCopy);

            // 6. Convert Button
            HWND hBtnConvert = CreateWindow("BUTTON", "CONVERT", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON | BS_DEFPUSHBUTTON, 
                         20, 210, 340, 40, hwnd, (HMENU)ID_BTN_CONVERT, NULL, NULL);
            // Make this font bold?
            HFONT hFontBold = CreateFont(18, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE, ANSI_CHARSET, 
                               OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, 
                               DEFAULT_PITCH | FF_SWISS, "Segoe UI");
            SendMessage(hBtnConvert, WM_SETFONT, (WPARAM)hFontBold, TRUE);

            break;
        }
        case WM_COMMAND: {
            if (LOWORD(wParam) == ID_BTN_SELECT) {
                SelectFile(hwnd);
            } else if (LOWORD(wParam) == ID_BTN_CONVERT) {
                ConvertFile(hwnd);
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

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    WNDCLASSEX wc;
    HWND hwnd;
    MSG Msg;

    wc.cbSize = sizeof(WNDCLASSEX);
    wc.style = 0;
    wc.lpfnWndProc = WndProc;
    wc.cbClsExtra = 0;
    wc.cbWndExtra = 0;
    wc.hInstance = hInstance;
    wc.hIcon = LoadIcon(NULL, IDI_APPLICATION);
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW); // White background
    wc.lpszMenuName = NULL;
    wc.lpszClassName = "UniversalConverterClass";
    wc.hIconSm = LoadIcon(NULL, IDI_APPLICATION);

    if (!RegisterClassEx(&wc)) {
        MessageBox(NULL, "Window Registration Failed!", "Error", MB_ICONEXCLAMATION | MB_OK);
        return 0;
    }

    hwnd = CreateWindowEx(
        0, "UniversalConverterClass", "Universal File Converter",
        WS_OVERLAPPEDWINDOW & ~WS_MAXIMIZEBOX & ~WS_THICKFRAME, // Fixed size
        CW_USEDEFAULT, CW_USEDEFAULT, 400, 310,
        NULL, NULL, hInstance, NULL);

    if (hwnd == NULL) {
        MessageBox(NULL, "Window Creation Failed!", "Error", MB_ICONEXCLAMATION | MB_OK);
        return 0;
    }

    ShowWindow(hwnd, nCmdShow);
    UpdateWindow(hwnd);

    while (GetMessage(&Msg, NULL, 0, 0) > 0) {
        TranslateMessage(&Msg);
        DispatchMessage(&Msg);
    }
    return Msg.wParam;
}
