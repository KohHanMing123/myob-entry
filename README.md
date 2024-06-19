# What the program does

Fill in MYOB Form from
![image](https://github.com/MKSanic/myob-entry/assets/76942402/eb4193a0-36a6-4b6f-847a-8afa90059dbe)
to
![image](https://github.com/MKSanic/myob-entry/assets/76942402/fef9efc9-163e-42a5-8f62-48359605545f)

using the values provided in data.xlsx:
![image](https://github.com/MKSanic/myob-entry/assets/76942402/10423797-92ac-499e-a4d6-e9932d8648e2)

# Notes
- Program does not read anything beyond column I, so feel free to make notes there

- Ex-Clients is the same as Ex-Adviser, so use Ex-Clients even if the adviser is Ex-Adviser

- If the entry requires 1 line, only fill Value 1 and Code 1 if entry requires 2 lines fill both Value 1 and 2, and Code 1 and 2

- Codes have to be a 5 digit number, if the code is 4-910x, put 49100. 4-01xx -> 40100

- Every adviser has a specific number from 1-20, and its added on to the specific code, so find the type of code you want and find the first adviser, IM and minus 1 from the code, which is what you will be using

- Fatal errors cause the program to not run until the error is amended in the file. They are caused by:

  - Invalid invoice date: must be 8 characters

  - Invalid Adviser, must be in the list on the right
 
  - Codes not being 5 digit
 
  - Missing Code or value

- Warnings can be ignored but requires the user to press enter for each warning. They are caused by:

  - Unrecognised providers

  - Unrecognised tax
 
  - Blank description
 


# How to install
  
1. Download [Python](https://www.python.org/) from the downloads section

2. Run the installer, tick 'Add python.exe to PATH', then 'Install Now'

  ![image](https://github.com/MKSanic/myob-entry/assets/76942402/45740e58-e658-475d-86b4-58d1d9de7c25)

3. Open cmd by typing 'cmd' in the windows search bar

  ![image](https://github.com/MKSanic/myob-entry/assets/76942402/2ca36e12-c17b-47aa-a693-0853d62279e2)

4. Run this command there `pip install pyautogui; pip install openpyxl`

  ![image](https://github.com/MKSanic/myob-entry/assets/76942402/9de376fd-8abd-486e-b95c-90d90c065da8)

5. Download the [code and template excel sheet](https://github.com/MKSanic/myob-entry/archive/refs/heads/main.zip)

> [!TIP]
> Move the program and template to Desktop to make it easier to access



# How to use

1. Fill in the blank.xlsx with the data

2. Rename the file to

 > data

> [!IMPORTANT]
> Make sure the excel file is in the same folder as the program etc. Desktop

3. Prepare MYOB entry page, ensuring it is maximised

   ![image](https://github.com/MKSanic/myob-entry/assets/76942402/a51bd330-0fcb-47b3-b4da-3227e1e2fe01)

 > [!CAUTION]
 > Make sure you have selected whether it is tax-inclusive or not. this will apply to all entries in the excel sheet
   
5. Run the program by double clicking it

6. Follow instructions on program

7. Once you get to this page, switch to MYOB quickly after it starts counting down

8. Done!!


