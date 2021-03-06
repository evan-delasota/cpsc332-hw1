CREATE TABLE EMPLOYEE (
    Fname VARCHAR(15) NOT NULL,
    Minit CHAR(1) NULL,
    Lname VARCHAR(15) NOT NULL,
    Ssn CHAR(9) NOT NULL,
    Bdate DATE NULL,
    Address VARCHAR(30) NULL,
    Sex CHAR(1) NULL,
    Salary DECIMAL(10, 2) NULL,
    Super_ssn CHAR(9) NULL,
    Dno INT NOT NULL,
    CONSTRAINT PK_EMPLOYEE PRIMARY KEY (Ssn),
    CONSTRAINT FK_EMPLOYEE_EMPLOYEE FOREIGN KEY (Super_ssn) REFERENCES EMPLOYEE (Ssn)
);

CREATE TABLE DEPARTMENT (
    Dname VARCHAR(15) NOT NULL,
    Dnumber INT NOT NULL,
    Mgr_ssn CHAR(9) NOT NULL,
    Mgr_start_date DATE NOT NULL,
    CONSTRAINT PK_DEPARTMENT PRIMARY KEY (Dnumber),
    CONSTRAINT FK_DEPARTMENT_EMPLOYEE FOREIGN KEY (Mgr_ssn) REFERENCES EMPLOYEE (Ssn)
);

ALTER TABLE EMPLOYEE ADD CONSTRAINT FK_DEPARTMENT FOREIGN KEY (Dno) REFERENCES DEPARTMENT (Dnumber);

CREATE TABLE DEPT_LOCATIONS (
    Dnumber INT NOT NULL,
    Dlocation VARCHAR(15) NOT NULL,
    CONSTRAINT PK_DEPT_LOCATIONS PRIMARY KEY (Dnumber, Dlocation),
    CONSTRAINT FK_DEPT_LOCATIONS_DEPARTMENT FOREIGN KEY (Dnumber) REFERENCES DEPARTMENT (Dnumber)
);

CREATE TABLE PROJECT (
    Pname VARCHAR(15) NOT NULL,
    Pnumber INT NOT NULL,
    Plocation VARCHAR(15) NULL,
    Dnum INT NOT NULL,
    CONSTRAINT PK_PROJECT PRIMARY KEY (Pnumber),
    CONSTRAINT UK_PROJECT UNIQUE (Pname),
    CONSTRAINT FK_PROJECT_DEPARTMENT FOREIGN KEY (Dnum) REFERENCES DEPARTMENT (Dnumber)
);

CREATE TABLE WORKS_ON (
    Essn CHAR(9) NOT NULL,
    Pno INT NOT NULL,
    Hours DECIMAL(3, 1) NULL,
    CONSTRAINT PK_WORKS_ON PRIMARY KEY (Essn, Pno),
    CONSTRAINT FK_WORKS_ON_EMPLOYEE FOREIGN KEY (Essn) REFERENCES EMPLOYEE (Ssn),
    CONSTRAINT FK_WORKS_ON_PROJECT FOREIGN KEY (Pno) REFERENCES PROJECT (Pnumber)
);

CREATE TABLE DEPENDENT (
    Essn CHAR(9) NOT NULL,
    Dependent_name VARCHAR(15) NOT NULL,
    Sex CHAR(1) NULL,
    Bdate DATE NULL,
    Relationship VARCHAR(8) NULL,
    CONSTRAINT PK_DEPENDENT PRIMARY KEY (Essn, Dependent_name),
    CONSTRAINT FK_DEPENDENT_EMPLOYEE FOREIGN KEY (Essn) REFERENCES EMPLOYEE (Ssn)
);

ALTER TABLE EMPLOYEE DROP CONSTRAINT FK_DEPARTMENT;

INSERT INTO EMPLOYEE (Fname, Minit, Lname, Ssn, Bdate, Address, Sex, Salary, Super_ssn, Dno) VALUES
    ('James', 'E', 'Borg', '888665555', '1937-11-10', '450 Stone, Houston, TX', 'M', 55000, NULL, 1),
    ('Franklin', 'T', 'Wong', '333445555', '1955-12-08', '638 Voss, Houston, TX', 'M', 40000, '888665555', 5),
    ('Jennifer', 'S', 'Wallace', '987654321', '1941-06-20', '291 Berry, Bellaire, TX', 'F', 43000, '888665555', 4),
    ('Alicia', 'J', 'Zelaya', '999887777', '1968-01-19', '3321 Castle, Spring, TX', 'F', 25000, '987654321', 4),
    ('Ahmad', 'V', 'Jabbar', '987987987', '1969-03-29', '980 Dallas, Houston, TX', 'M', 25000, '987654321', 4),
    ('John', 'B', 'Smith', '123456789', '1965-01-09', '731 Fonden, Houston, TX', 'M', 30000, '333445555', 5),
    ('Ramesh', 'K', 'Narayan', '666884444', '1962-09-15', '975 Fire Oak, Humble, TX', 'M', 38000, '333445555', 5),
    ('Joyce', 'A', 'English', '453453453', '1972-07-31', '5631 Rice, Houston, TX', 'F', 25000, '333445555', 5);

INSERT INTO DEPARTMENT (Dname, Dnumber, Mgr_ssn, Mgr_start_date) VALUES
    ('Research', 5, '333445555', '1988-05-22'),
    ('Administration', 4, '987654321', '1995-01-01'),
    ('Headquarters', 1, '888665555', '1981-06-19');

ALTER TABLE EMPLOYEE ADD CONSTRAINT FK_DEPARTMENT FOREIGN KEY (Dno) REFERENCES DEPARTMENT (Dnumber);

INSERT INTO DEPT_LOCATIONS (Dnumber, Dlocation) VALUES
    (1, 'Houston'),
    (4, 'Stafford'),
    (5, 'Bellaire'),
    (5, 'Sugarland'),
    (5, 'Houston');

INSERT INTO PROJECT (Pname, Pnumber, Plocation, Dnum) VALUES
    ('ProductX', 1, 'Bellaire', 5),
    ('ProductY', 2, 'Sugarland', 5),
    ('ProductZ', 3, 'Houston', 5),
    ('Computerization', 10, 'Stafford', 4),
    ('Reorganization', 20, 'Houston', 1),
    ('Newbenefits', 30, 'Stafford', 4);

INSERT INTO WORKS_ON (Essn, Pno, Hours) VALUES
    ('123456789', 1, 32.5),
    ('123456789', 2, 7.5),
    ('666884444', 3, 40.0),
    ('453453453', 1, 20.0),
    ('453453453', 2, 20.0),
    ('333445555', 2, 10.0),
    ('333445555', 3, 10.0),
    ('333445555', 10, 10.0),
    ('333445555', 20, 10.0),
    ('999887777', 30, 30.0),
    ('999887777', 10, 10.0),
    ('987987987', 10, 35.0),
    ('987987987', 30, 5.0),
    ('987654321', 30, 25.0),
    ('987654321', 20, 15.0),
    ('888665555', 20, NULL);

INSERT INTO DEPENDENT (Essn, Dependent_name, Sex, Bdate, Relationship) VALUES
    ('333445555', 'Alice', 'F', '1986-04-05', 'Daughter'),
    ('333445555', 'Theodore', 'M', '1983-10-25', 'Son'),
    ('333445555', 'Joy', 'F', '1958-05-03', 'Spouse'),
    ('987987987', 'Abner', 'M', '1942-02-28', 'Spouse'),
    ('123456789', 'Michael', 'M', '1988-01-04', 'Son'),
    ('123456789', 'Alice', 'F', '1988-12-30', 'Daughter'),
    ('123456789', 'Elizabeth', 'F', '1967-05-05', 'Spouse');
GO
