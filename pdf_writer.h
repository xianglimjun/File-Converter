#ifndef PDF_WRITER_H
#define PDF_WRITER_H

#include <stdio.h>
#include <stdlib.h>
#include <string.h>

/* 
 * Minimal PDF Writer for Single Image
 * Supports embedding raw pixel data (RGB) into a PDF.
 * Note: Real-world PDF writers are much more complex. This is a minimal implementation
 * to satisfy the "no dependency" requirement for the specific task of Image->PDF.
 */

// Helper to write a number as a string
static void write_int(FILE* f, int i) {
    fprintf(f, "%d", i);
}

// Writes a basic PDF file wrapping the RGB image data
// Returns 1 on success, 0 on failure
int create_pdf_from_image(const char* filename, int width, int height, int channels, unsigned char* data) {
    if (channels != 3) {
        // For simplicity, this minimal writer insists on RGB. 
        // If detection shows RGBA, we might need to strip alpha or just write it but PDF expects /DeviceRGB usually.
        // We will handle 4 channels by dropping alpha in the loop below if needed.
    }

    FILE* f = fopen(filename, "wb");
    if (!f) return 0;

    // PDF Header
    const char* header = "%PDF-1.4\n";
    fwrite(header, 1, strlen(header), f);

    // We need to keep track of byte offsets for the XREF table
    long offsets[6]; // We will have 5 objects + 1 dummy (index 0)

    // Object 1: Catalog
    offsets[1] = ftell(f);
    fprintf(f, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");

    // Object 2: Pages
    offsets[2] = ftell(f);
    fprintf(f, "2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n");

    // Object 3: Page
    // MediaBox is [0 0 width height]
    offsets[3] = ftell(f);
    fprintf(f, "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 %d %d] /Resources << /XObject << /Img4 4 0 R >> >> /Contents 5 0 R >>\nendobj\n", width, height);

    // Object 4: Image XObject
    offsets[4] = ftell(f);
    fprintf(f, "4 0 obj\n<< /Type /XObject /Subtype /Image /Width %d /Height %d /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length %d >>\nstream\n", width, height, width * height * 3);
    
    // Write image data
    // Data comes in as RGB or RGBA from stbi_load. We need straight RGB.
    long stream_start = ftell(f);
    if (channels == 3) {
        fwrite(data, 1, width * height * 3, f);
    } else if (channels == 4) {
        for (int i = 0; i < width * height; i++) {
            fwrite(&data[i * 4], 1, 3, f); // Write RGB, skip A
        }
    } else if (channels == 1) {
         // Grayscale -> RGB
         for (int i = 0; i < width * height; i++) {
            unsigned char g = data[i];
            fwrite(&g, 1, 1, f);
            fwrite(&g, 1, 1, f);
            fwrite(&g, 1, 1, f);
        }
    }
    
    fprintf(f, "\nendstream\nendobj\n");

    // Object 5: Content Stream (Draw the image)
    offsets[5] = ftell(f);
    // q = save state, %d %d 0 0 re = rectangle, W n = clip (not strictly needed for just drawing image)
    // cm = current transformation matrix. To fill the page: width 0 0 height 0 0 cm
    // Do = paint copy of XObject
    char content[256];
    snprintf(content, sizeof(content), "q %d 0 0 %d 0 0 cm /Img4 Do Q", width, height);
    fprintf(f, "5 0 obj\n<< /Length %llu >>\nstream\n%s\nendstream\nendobj\n", (unsigned long long)strlen(content), content);

    // XREF
    long startxref = ftell(f);
    fprintf(f, "xref\n0 6\n0000000000 65535 f \n");
    for (int i = 1; i <= 5; i++) {
        fprintf(f, "%010ld 00000 n \n", offsets[i]);
    }

    // Trailer
    fprintf(f, "trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n%ld\n%%%%EOF", startxref);

    fclose(f);
    return 1;
}

#endif
