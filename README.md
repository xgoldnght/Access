# SQL Server
```sql
-- Создание базы данных
CREATE DATABASE HotelSystem;
GO
```

```sql
-- Использование базы данных
USE HotelSystem;
GO
```

```sql
-- Таблица пользователей
CREATE TABLE users (
    id_user INT IDENTITY(1,1) PRIMARY KEY,
    login NVARCHAR(50),
    password NVARCHAR(255),
    access_level NVARCHAR(50),
    login_attempts INT,
    is_blocked BIT DEFAULT 0,
    last_auth_date DATE DEFAULT GETDATE(),
    account_confirmed BIT DEFAULT 0
);
```

```sql
-- Таблица сотрудников
CREATE TABLE employees (
    id_employee INT IDENTITY(1,1) PRIMARY KEY,
    full_name NVARCHAR(100),
    phone NVARCHAR(20),
    id_user INT FOREIGN KEY REFERENCES users(id_user),
    hire_date DATE,
    dismissal_date DATE
);
```

```sql
-- Таблица клиентов
CREATE TABLE clients (
    id_client INT IDENTITY(1,1) PRIMARY KEY,
    full_name NVARCHAR(100),
    phone NVARCHAR(20),
    id_user INT FOREIGN KEY REFERENCES users(id_user)
);
```

```sql
-- Таблица номеров
CREATE TABLE rooms (
    id_room INT IDENTITY(1,1) PRIMARY KEY,
    floor INT,
    room_number NVARCHAR(10),
    category NVARCHAR(50),
    status NVARCHAR(50)
);
```

```sql
-- Таблица оплат
CREATE TABLE payments (
    id_payment INT IDENTITY(1,1) PRIMARY KEY,
    id_client INT FOREIGN KEY REFERENCES clients(id_client),
    price DECIMAL(10, 2),
    payment_date DATE
);
```

```sql
-- Таблица бронирований
CREATE TABLE bookings (
    id_booking INT IDENTITY(1,1) PRIMARY KEY,
    id_client INT FOREIGN KEY REFERENCES clients(id_client),
    id_room INT FOREIGN KEY REFERENCES rooms(id_room),
    check_in_date DATE,
    check_out_date DATE
);
```











