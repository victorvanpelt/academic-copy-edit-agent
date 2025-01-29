## Academic Copy Editor Agent

This is an academic correcting agent that copy edits a working paper and saves the updated document, including a track-changes copy. It runs through every paragraph, correcting grammar and spelling while also incrementally improving the style. It also tries to leave the paragraph structure, content, formatting, and terminology as it is.

## How to use?

1.  Add your OpenAI API key as an environmental variable (it is called in line 8 of "correct_paper.py")

2.  OPTIONAL: Adjust instructions in lines 22-33 and line 56 of "correct_paper.py"

3.  Save your paper as a "paper.docx" in the "0_input" folder. Ensure it includes the headings: "Abstract," "Introduction," and "References." The script will use these headings as reference points.

4.  Run the python file "correct_paper.py" It may take a while, so grab a coffee. The script will print its progress (e.g., "Processed paragraph 2/X" etc) 

5.  When the code finishes, you can enjoy your "free" clean copy edit under "1_output/edited_paper.docx" and track-changes copy edit under "1_output/trackchanges_paper.docx"
