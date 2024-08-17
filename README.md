# Price Offer Opulent Fabricator

A finance management application for Multimedica ScO

## Table of Contents

- [About](#about)
- [Getting Started](#getting-started)
  - [Prerequisites](#prerequisites)
  - [Installation](#installation)
  - [Features](#features)
- [Usage](#usage)

  
## About

POOF is a finance management program designed for the internal use of a medical equipment company based in Egypt called Multimedica ScO. 
This project started when I suggested automating the price offer/quotation-making process from filling data out in a Microsoft Excel data sheet and stylizing it manually every time to a web application.
<br>
More features might be added as I go as this is a work in progress, and I will keep this repo updated with any extra features added later on

## Getting Started


### Prerequisites

Python, 
Werkzeug.security,
Flask,
wtforms,
sqlalchemy,
flask_wtf,
pandas,
openpyxl,
flask_session,
PIL,
qrcode,
datetime,
os,
pythoncom,
win32com.client,
math

### Features

- This application takes input from the office secretary, fills out a premade, stylized form, and then exports it into an Excel template, making the process more streamlined.
- This application also keeps track of each submitted price offer and which user submitted it, allowing for better performance monitoring at the company.
- Each price offer comes equipped with a QR code that allows the end users to later scan the code and get directed to a digital version of the printed price offer they got in hand, the digital version also includes pictures of the products.
- The app allows for the editing of the pre-existing price list, either via editing a single product or adding a percentage to increase the prices of all products.
- Multiple levels of authorization and access are made to ensure that data would not be tampered with. 


### Installation

This project is still a work in progress, an installation guide will be provided as soon as the project is fully deployed.

### Usage

This project is still a work in progress, a user manual will be provided as soon as the project is fully deployed.
