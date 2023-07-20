# Excel-Based-Employee-Management-Python-CRUD-Application
In this project, i have developed a CRUD (Create, Read, Update, Delete) application for managing employee data using Excel as the database.

This version represents an upgrade from our previous project (Capstone_Project_Modul1), in this version i implemented a database to store and track data changes in real-time. Now, users can easily add, update, and delete data, with all modifications instantly applied to the Excel file.

One of the most significant improvements is the employee ID generator. It now ensures that once an employee ID is used and subsequently deleted, it becomes permanently unavailable for future use. The deleted ID is stored in a separate file called 'Used ID', and if a similar ID is generated in the future, the system will automatically assign a new unique ID. Furthermore, updating employee data will no longer affect the employee ID, providing a more stable identifier.

To achieve a cleaner and less repetitive codebase, i have introduced additional functions. These include a function for displaying menus and another for printing error messages. By using *args as a parameter in the function, i can easily print the menu content and making format for printing error messages consistently, improving code readability and maintenance.

Please note that while i use personal information for generating ID, it is generally advisable to use more secure and privacy-conscious methods in real-world applications. The ID generator that i make is just a demonstration of making a function that can generate a unique identifier.


