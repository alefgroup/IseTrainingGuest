# IseTrainingGuest

## Overview
The application is used to create guest user access within Cisco ISE. These users are defined in source data table. 
These users access is created by calling od Cisco ISE REST API. The user accesses are created using Cisco ISE REST API calls. 
Application prints report with these created accesses.
Application is used as a sponsor portal, where candidates for training need create accounts in Cisco ISE.
These candidates are loaded from database table Users (each candidate has record in table row).
Application create guest account for each candidate in Cisco ISE with time restricted validity. This gest account is created by calling Cisco API.
At the end application creates report with candidates' accesses a sends it via email to preconfigured recipient.

## Installation
This is a console application. This means that it is necessary to build the .NET application to start. Next, you need to configure the application in the configuration file AppData\Config.xml 

Prerequisities:
-install MS SQL Express
-create database table Users (use DB script users.sql)
-modify config.xml (set connection string to MS SQL database)
-set ISE setting base URL, sponsor portal

## Usage
It is a console application. The binary file obtained by the build is run directly from the console.
