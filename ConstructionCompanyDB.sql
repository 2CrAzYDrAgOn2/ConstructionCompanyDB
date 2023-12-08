CREATE DATABASE ConstructionCompanyDB;
USE ConstructionCompanyDB;

CREATE TABLE Projects (
    ProjectID INT PRIMARY KEY IDENTITY(1,1),
    ProjectName NVARCHAR(255) NOT NULL,
    StartDate DATE,
    EndDate DATE,
    Budget INT,
    Status NVARCHAR(50)
);

CREATE TABLE Customers (
    CustomerID INT PRIMARY KEY IDENTITY(1,1),
    CustomerName NVARCHAR(255) NOT NULL,
    ContactPerson NVARCHAR(255),
    ContactNumber NVARCHAR(20),
    Email NVARCHAR(255)
);

CREATE TABLE Employees (
    EmployeeID INT PRIMARY KEY IDENTITY(1,1),
    FirstName NVARCHAR(50) NOT NULL,
    LastName NVARCHAR(50) NOT NULL,
    Position NVARCHAR(50),
    HireDate DATE,
    Salary INT,
    Email NVARCHAR(255),
    PhoneNumber NVARCHAR(20)
);

CREATE TABLE Materials (
    MaterialID INT PRIMARY KEY IDENTITY(1,1),
    MaterialName NVARCHAR(255) NOT NULL,
    UnitPrice INT,
    QuantityInStock INT
);

CREATE TABLE ProjectMaterials (
    ProjectID INT,
    MaterialID INT,
    QuantityUsed INT,
    PRIMARY KEY (ProjectID, MaterialID),
    FOREIGN KEY (ProjectID) REFERENCES Projects(ProjectID),
    FOREIGN KEY (MaterialID) REFERENCES Materials(MaterialID)
);

CREATE TABLE Registration (
	UserID INT PRIMARY KEY IDENTITY(1,1),
	UserLogin VARCHAR(50),
	UserPassword VARCHAR(50),
	IsAdmin bit
);

INSERT INTO Projects (ProjectName, StartDate, EndDate, Budget, Status)
VALUES 
    ('Building A', '2023-01-01', '2023-06-30', 500000, 'In Progress'),
    ('Renovation B', '2023-03-15', '2023-08-31', 300000, 'Not Started'),
    ('Infrastructure C', '2023-02-01', '2023-12-31', 800000, 'Completed');

INSERT INTO Customers (CustomerName, ContactPerson, ContactNumber, Email)
VALUES 
    ('ABC Construction', 'John Smith', '123-456-7890', 'john@abcconstruction.com'),
    ('XYZ Builders', 'Jane Doe', '987-654-3210', 'jane@xyzbuilders.com');

INSERT INTO Employees (FirstName, LastName, Position, HireDate, Salary, Email, PhoneNumber)
VALUES 
    ('Michael', 'Johnson', 'Project Manager', '2022-05-01', 75000, 'michael@email.com', '555-1234'),
    ('Emily', 'Williams', 'Site Engineer', '2022-06-15', 60000, 'emily@email.com', '555-5678');

INSERT INTO Materials (MaterialName, UnitPrice, QuantityInStock)
VALUES 
    ('Concrete', 100, 500),
    ('Steel', 150, 300),
    ('Bricks', 50, 1000);

INSERT INTO ProjectMaterials (ProjectID, MaterialID, QuantityUsed)
VALUES 
    (1, 1, 200),
    (1, 2, 100),
    (2, 3, 500),
    (3, 1, 300),
    (3, 2, 200);

INSERT INTO Registration (UserLogin, UserPassword, IsAdmin)
VALUES
	('admin', 'admin', 1),
	('user', 'user', 0);

SELECT * FROM Projects;
SELECT * FROM Customers;
SELECT * FROM Employees;
SELECT * FROM Materials;
SELECT * FROM ProjectMaterials;
SELECT * FROM Registration;

DROP DATABASE ConstructionCompanyDB;