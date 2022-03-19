# vba-ip-validation
![GitHub top language](https://img.shields.io/github/languages/top/tankalxat34/vba-ip-validation)
![skill](https://img.shields.io/badge/Microsoft%20Excel%20VBA-107C41?logo=microsoft&logoColor=white)
![GitHub](https://img.shields.io/github/license/tankalxat34/vba-ip-validation?logo=github&logoColor=white)

[![vbaRequests](https://img.shields.io/badge/module-vba%20Requests-5C2D91?logo=.net&logoColor=white&style=for-the-badge)](https://github.com/tankalxat34/vbaRequests)

This repository is a example for my **[vbaRequests module](https://github.com/tankalxat34/vbaRequests)**. If you want to use IP validation code, you need to install [vbaRequests](https://github.com/tankalxat34/vbaRequests).

This is a simple code for automatically verification user's IP-address that can help you to block all users who has you Excel Book and who should not have access to it (that is, these users will not be able to open it).

# Install
You need to install [`validationIP.bas`](https://github.com/tankalxat34/vba-ip-validation/blob/main/validationIP.bas) from `main` branch and place into your book.

# How it work?
A user whose IP address is not saved to a file with the addresses of users who have access to your book opens the book containing this code and receives an error that his IP address is not saved on the server. Then the book closes. Thus, an unidentified user will not be able to view your Excel workbook and change anything there.

If the user's IP is contained in a file with IP addresses, nothing happens and the book opens correctly without giving any access errors.
