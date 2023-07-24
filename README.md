# Excel-Based-Employee-Management-Python-CRUD-Application
In this project, i have developed a CRUD (Create, Read, Update, Delete) application for managing employee data using Excel as the database.

This version represents an upgrade from my previous project (Capstone_Project_Modul1), in this version i implemented a database to store and track data changes in real time. Now, users can easily add, update, and delete data, with all modifications instantly applied to the Excel file.

One of the most significant improvements is the the addition of database to store changes and new data. Additionally, for the employee ID generator, now once an employee ID is used and subsequently deleted, it becomes permanently unavailable for future use. The deleted ID is stored in a separate file called 'Used ID', and if a similar ID is generated in the future, the system will automatically assign a new unique ID. Furthermore, updating employee data will no longer affect the employee ID, providing a more stable identifier.

To achieve a cleaner and less repetitive codebase, i have introduced additional functions. These include a function for displaying menus and another for printing error messages. By using *args as a parameter in the function, i can easily print the menu content and making format for printing error messages consistently, improving code readability and maintenance.

Please note that while i use personal information for generating employee ID, it is generally advisable to use more secure and privacy-conscious methods in real world applications. The ID generator that i make is just a demonstration of making a function that can generate a unique identifier.

---
To use this you just have to download all the file, then in the code change the value of 'employee_path', 'dept_path', and 'usedID_path'


To use this application, follow these simple steps:
  1. Download all the necessary files.
  2. Open the code and navigate to the section where the file paths are defined. Locate the variables: 'employee_path', 'dept_path', and 'usedID_path' in       the 'Main_flow' function.
  3. Modify the values of 'employee_path', 'dept_path', and 'usedID_path' to the appropriate file paths on your local machine where you want to store the       employee data, department data, and used employee IDs, respectively.
  4. Save the changes to the code.
  5. Run the program to execute the Employee Management CRUD application.

With these modifications, the application will use the specified file paths for managing the employee data and used employee IDs, enabling you to interact with the Employee Management system using your preferred storage locations.
