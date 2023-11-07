A simple Cover Letter generator. 
<br>
<br>
It uses a template file to generate a cover letter based on the information you provide.
<br>
<br>
The idea is to ease the just replace the company name and position name for each position that you apply to, since it is practically impossible to write a cover letter for each job that you want to apply to. Sadly, the current job market is a numbers game more than anything else, so here we are.
<br>
<br>
The requirements.txt file is there if you need it. Simply put, you only need docx, python-docx, re and pysimplegui to be installed on your system.
<br>
<br>
Once you have this script (I use Windows), create a shortcut of the "cover_letter_script.py" file by right-clicking on it. Right click on the shortcut and select "Properties". In the "Shortcut Key" field, hold whatever combination of keys you want to use to run the script. I use "Ctrl + Alt + C" for example.
<br>
<br>
Feel free to use, redistribute, modify, etc. this script. I hope it helps you in your job search. I hope it helps my job search too :D.
<br>
<br>
I am attaching a sample of the folder structure that I use for this project so that everyone else can also use it:
<br>
cover_letter_editor/
├── generated_cover_letters/
│   ├── .gitkeep
│   ├── generated_output_1.docx
│   ├── generated_output_2.docx
│   ├── ...
│   └── generated_output_n.docx
│
├── template_cover_letter/
│   ├── .gitkeep
│   └── YOUR_TEMPLATE_COVER_LETTER.docx
│
├── .gitignore
├── cover_letter_script.ipynb
├── cover_letter_script.py
├── README.md
└── requirements.txt