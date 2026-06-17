# HTML to PowerPoint Converter

A local Python desktop application with a graphical interface that converts HTML files into widescreen (16:9) PowerPoint presentations by taking high-resolution browser screenshots and organizing them into slides.

## Features

- **High-Resolution Output:** Emulates a high-DPI (Retina-equivalent) display to render sharp text and graphics.
- **Optimized File Size:** Automatically compresses and saves slides in JPEG format (85% quality) to prevent presentations from becoming excessively large.
- **Flexible Modes:**
  - **Vertical Scroll:** Captures standard web documents page-by-page.
  - **Section Tags:** Automatically isolates and captures separate `<section>` elements (useful for HTML-based slide decks).
- **Aspect-Ratio Preservation:** Automatically scales and centers content on widescreen slides without stretching or distortion.
- **Simple GUI:** Tkinter-based interface for easy file selection.

## Prerequisites

Make sure you have Python installed. You will need to install the following dependencies:

```bash
# Install required Python libraries
pip install playwright python-pptx Pillow

# Download the browser engine used for screenshots
playwright install chromium