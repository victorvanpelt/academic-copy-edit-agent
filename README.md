## Academic Copy Editor Agent

This is a script that copy edits academic papers and saves the updated document, including a separate track-changes copy in docx format. It runs through every paragraph, correcting exclusively grammar, spelling, and style. It also tries to leave the paragraph structure, substance, formatting, and terminology intact.

## Requirements
- You will need to be able to run .py files and have an python environment, while importing the following packages: openai, os, docx, win32com.client, and re.
- You will also need an OPENAI account including an API key allowing you to connect the agent to open AI.
- The document you want to edit needs to be in a docx format, ideally without figures, appendices, and tables. 
- The document should also include, at the very least, the following headers in the following order: Abstract, Introduction, References. The script looks for these headers to use them as reference points.

## How to use?

1.  Add your OpenAI API key as an environmental variable (it is called in line 25 of "correct_paper.py")

2.  OPTIONAL: Adjust instructions in lines 98 of "correct_paper.py"

3.  OPTIONAL: Adjust the model you'd like to use on line 13 of "correct_paper.py" GPT models work better with the specific instructions.

4.  Save your paper as a "paper.docx" in the "0_input" folder. Ensure it includes the headings: "Abstract," "Introduction," and "References." The script will use these headings as reference points.

5.  Run the python file "correct_paper.py" It may take a while, so grab a coffee. The script will print its progress (e.g., "Processed paragraph 2/X" etc) 

6.  When the code finishes, you can enjoy your "free" clean copy edit under "1_output/edited_paper.docx" and track-changes copy edit under "1_output/trackchanges_paper.docx"

## Known issues
- The script cannot handle and thus automatically deletes footnotes. Just reject these changes in the track changes document.
- The script does not interact well with word reference managers. This may create unnecessary trackchanges in the trackchanges_paper.docx when comparing.