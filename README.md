# pdf2word
PDF to Word Converter

The .exe file is included in the dist for windows devices. 

This project is a lightweight document management utility that makes it easy to open/edit Word documents, manage multiple PDFs as appendices, reorder them visually, and convert/export the final output in a clean and intuitive interface.

Key Features

-> Word File Handling

    Open existing Word documents or create new ones.

    Status indicator with a colored dot (●) and a descriptive label shows whether a Word file is active.

-> PDF Appendix Management

    Select multiple PDF files to add as appendices.

    Each PDF is displayed as a stack card with:

      File name and page count.

      Buttons for Move Up, Move Down, and Delete, with hover effects and spacing for a clean look.

    Users can reorder PDFs interactively, ensuring the final document respects their intended order.

-> Conversion Workflow

    Once configured, the Convert button processes the active Word document and integrates the selected PDFs.

    A dialog box confirmation informs the user when the conversion is complete.

    A Clear button allows resetting all appendices and status labels.

-> Polished GUI Design

    Minimalistic, card-like stack frames with soft gray borders.

    Hover effects for better interactivity.

    Buttons aligned neatly at the center with consistent sizing.


For mac devices, to generate the exe file, follow the below steps:
  => Clone the repo
  => pip install pyinstaller
  => pyinstaller --onefile --noconsole --clean --collect-all docxcompose .\res2.py 
  => You’ll get a .app (mac application bundle) or a binary in dist/.
  

