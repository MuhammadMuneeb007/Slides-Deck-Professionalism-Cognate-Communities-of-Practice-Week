 
# Setup Instructions

## 1. Download Anaconda

Download Anaconda from the following link:

[Download Anaconda](https://www.anaconda.com/download)

## 2. Create and Activate Conda Environment

Run the following commands in your terminal to create and activate a new conda environment:

```bash
conda create -n academictools
conda activate academictools
```

## 3. Install Required Packages

Install the necessary Python packages using pip:

```bash
pip install python-pptx
pip install spellchecker
pip install pyspellchecker
pip install textstat
pip install pdf2docx PyPDF2
pip install docx2pdf
pip install imgkit
pip install enchant
pip install pyenchant
pip install LanguageTool
pip install language-tool-python
pip install gingerit
pip install deepgrammar
pip install deepgram-sdk
pip install grammar-check
pip install textblob
```
Here's the content formatted in Markdown:

 
## 4. Prepare Your Files

Create a folder named `Files` and place all your slides in it.

The directory structure should look like this:

```
C:.
│   Cognate Communities of Practice Week.pptx
│   Installation.txt
│
└───Demo1
    │   Approach1-UsingPython.py
    │   Approach2-UsingPythonAndWindowsAPI.py
    │
    └───Files
            TECH1100 Week 01 Workshop.pptx
            TECH1100 Week 02 Workshop.pptx
            TECH1100 Week 03 Workshop.pptx
            TECH1100 Week 04 Workshop.pptx
            TECH1100 Week 05 Workshop.pptx
            TECH1100 Week 06 Workshop.pptx
            TECH1100 Week 07 Workshop.pptx
            TECH1100 Week 08 Workshop.pptx
            TECH1100 Week 09 Workshop.pptx
            TECH1100 Week 10 Workshop.pptx
            TECH1100 Week 11 Workshop.pptx
            TECH1100 Week 12 Workshop.pptx
```

Ensure the files are placed correctly as shown in the structure above.
```
 
            
## 5. Execute Approaches

### Approach 1: Using Python

- **Speed:** Fast
- **Limitations:** Limited options (SmartArt not processed)

Run the following command:

```bash
python Approach1-UsingPython
```

### Approach 2: Using Python and Windows API

- **Speed:** Slow
- **Features:** Advanced features. Only processes words on Windows and requires COM library.

Run the following command:

```bash
python Approach2-UsingPythonAndWindowsAPI
```
```

You can save this content in a `.md` file and upload it to GitHub or include it in your project documentation.
