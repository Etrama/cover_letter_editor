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
The method of creating a shortcut doesn't work for me personally anymore since the shortcut opens the file in VS Code. I'm trying a workaround by instead running the python file as an app using [pyinstaller](https://pyinstaller.org/en/stable/).  

Essentially, you can now run the script even without having python installed on your system as long as you have the folder generated by pyinstaller and you run the executable. Please do not run any executable files that you do not trust.

```pyinstaller cover_letter_script.py```

This will create a build and dist folder. Use the executable called cover_letter_script.exe from the dist folder. 
<br>
<br>
That didn't work either. Trying batch files now. If you have python in your path just write the following line in a notepad file and save it as a .bat file. You can then run the script by double clicking on the .bat file or assigning a shortcut to it.

```python cover_letter_script.py```

This approach worked for me, as of this writing.
<br>
<br>
Feel free to use, redistribute, modify, etc. this script. I hope it helps you in your job search. I hope it helps my job search too :D.
<br>
<br>
I am attaching a sample of the folder structure that I use for this project so that everyone else can also use it:
<br>
cover_letter_editor/  
├── build/  
├── dist/  
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