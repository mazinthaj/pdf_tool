# pdf_tool
- A tool for merging and deleting pages of pdfs.
- page deletion formula: To delete the pages in doc1 and doc2 the format will be: [1: 2, 3, 5-8][2: 3, 5-20]
- This means it will delete pages: 2, 3, 5, 6, 7, 8 in doc1 and 3, 5 - 20 (inclusive) pages in doc2.
  

When the "Remove Blank Pages" is checked, it removes pages that are above 97% empty. You can modify this in the code (white_ratio > 0.97)

![image](https://github.com/user-attachments/assets/75ff19b0-daca-41aa-b284-b01680144ab0)

