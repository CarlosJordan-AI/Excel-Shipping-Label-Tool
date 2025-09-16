# Excel-Shipping-Label-Tool
Excel VBA tool for automated generation of shipping labels, Bills of Lading (BOL), and QR-coded logistics documents.

This tool was developed to simplify logistics workflows by combining Excel automation (VBA) with dynamic QR code generation hosted via GitHub Pages.



# Excel Shipping Label & Logistics Tool

This Excel VBA-based tool automates:
- Printing shipping labels per FDC (single & batch).
- Generating Bills of Lading (BOL).
- Exporting QR labels with dynamic placement.
- Automatic logo & carrier image insertion.

## Features
- Minimal user input required.
- Handles multi-FDC and handling units.
- QR code integration (works on Android & iPhone).
- Automated PDF export.
- Automated shipping label generation (single & batch).
- Automated Bill of Lading (BOL) export.
- Lookup formulas tied to a Base dataset for hardware, slab, frame, and kit configurations.

  

## 📂 Repository Contents
Excel-Shipping-Label-Tool/
│
├─ Modules/                      # Exported VBA macros
├─ Example Files/                # Updated example workbooks / PDFs (optional)
├─ docs/
│  └─ qr/                        # QR viewer (GitHub Pages)
│     ├─ index.html
│     ├─ logo/
│     └─ data/                   # Example order data used by the viewer
│        ├─ WG96895753.txt
│        └─ WN30452442.txt
├─ Shipping Labels and Bill Of Lading Tool SOP V3.docx
└─ Version3.0.xlsm               # Latest working version of the tool


## How to Use
1. Download the latest release (`Version3.0.xlsm`).
2. Enable macros in Excel.
3. Follow the [User Manual].

## Live QR Viewer
The tool encodes order details into QR codes that open directly in a hosted viewer.
👉 Hosted at:
https://carlosjordan-ai.github.io/Excel-Shipping-Label-Tool/qr/

## Example Outputs
You can try scanning these QR codes or open the URLs directly.

Here are live examples you can open directly in the hosted QR viewer:
- Customer Order WG96895753 → Open Example https://carlosjordan-ai.github.io/Excel-Shipping-Label-Tool/qr/?id=WG96895753
- Customer Order WN30452442 → Open Example https://carlosjordan-ai.github.io/Excel-Shipping-Label-Tool/qr/?id=WN30452442

(These examples are stored as .txt files in docs/qr/, and the QR viewer loads them dynamically.)


## License
MIT – free to use for demonstration purposes.
