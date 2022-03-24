# vba-ip-validation
<img src="https://raw.githubusercontent.com/tankalxat34/vba-ip-validation/readme_content/icon_word.svg" width="40px"/> <img src="https://raw.githubusercontent.com/tankalxat34/vba-ip-validation/readme_content/icon_excel.svg" width="40px"/> <img src="https://raw.githubusercontent.com/tankalxat34/vba-ip-validation/readme_content/icon_powerpoint.svg" width="40px"/>

![GitHub top language](https://img.shields.io/github/languages/top/tankalxat34/vba-ip-validation)
![GitHub](https://img.shields.io/github/license/tankalxat34/vba-ip-validation?logo=github&logoColor=white)
[![vbaRequests](https://img.shields.io/badge/core-vbaRequests-5C2D91?logoColor=white)](https://github.com/tankalxat34/vbaRequests)

This repository is a example for my **[vbaRequests module](https://github.com/tankalxat34/vbaRequests)**. If you want to use IP validation code, you need to install [vbaRequests](https://github.com/tankalxat34/vbaRequests).

This is a simple code for automatically verification user's IP-address that can help you to block all users who has you Excel Book and who should not have access to it (that is, these users will not be able to open it).

# Install
- Install the **[vbaRequests module](https://raw.githubusercontent.com/tankalxat34/vbaRequests/main/vbaRequests.bas)**.
- Import downloaded module in your Document, Book or Presentation.
- Install [`validationIP.bas`](https://raw.githubusercontent.com/tankalxat34/vba-ip-validation/main/validationIP.bas) from `main` branch.
- If you using **Microsoft Excel**:
  * Open **"`VBAProject` → `Microsoft Excel Objects` → `ThisBook`"** and paste here code from downloaded file [`validationIP.bas`](https://raw.githubusercontent.com/tankalxat34/vba-ip-validation/main/validationIP.bas).
- If you using **Microsoft Word**:
  * Open **"`VBAProject` → `Microsoft Word Objects` → `ThisDocument`"** and paste here code from downloaded file [`validationIP.bas`](https://raw.githubusercontent.com/tankalxat34/vba-ip-validation/main/validationIP.bas).

# How it work?
A user whose IP address is not saved to a file with the addresses of users who have access to your book opens the book containing this code and receives an error that his IP address is not saved on the server. Then the book closes. Thus, an unidentified user will not be able to view your Excel workbook and change anything there.

If the user's IP is contained in a file with IP addresses, nothing happens and the book opens correctly without giving any access errors.
