VB.NET-Desktop-App-Generator
============================

VB.NET Desktop App Generator for creating VB.NET Desktop applications - lets you quickly set up a project with sensible defaults and best practices.

## Project Structure

Overview

    ├── My Project/             - Configuration of the application
    ├── Classes/                - Contains Classes
    ├── Config/                 - Exe. Application configuration on runtime
    │   ├── config.ini/         - Contains configuration of the application
    ├── Dataset/                - Application Datasets
    ├── Forms/                  
    │   ├── Form1.vb            - Optional Form
    │   ├── frmConnector.vb     - Form used to build connection between application and database, connection will be stored on config.ini
    │   ├── frmLogin.vb         - Login Form
    │   ├── frmMain.vb          - This the MDIParent of the application, serve as the parent among all forms.
    │   ├── frmProperties       - Form used to display the database connection properties
    ├── Images/                 - Contains Images requires in the application
    ├── Modules/                - Contains all the Module files
    │   ├── modConnection.vb    - Contains code to initialize the app configuration, database connection and application execution
    │   ├── modEmployees.vb     - (Optional) Contains Controller Code for Employees
    │   ├── modFunctions.vb/    - Contains custom made functons and procedures for easy and faster development
    │   ├── modINIParser.vb/    - This parses the values in a particular .INI file
    ├── Reports/                - Contains Reports
    │   ├── rtpSample.rpt       - include Sample reports
    ├── App.config              - Application configuration
