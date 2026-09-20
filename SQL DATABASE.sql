CREATE DATABASE transaction_validator;
SHOW DATABASES;

USE transaction_validator;
SELECT user, host
FROM mysql.user;

USE transaction_validator;
SELECT * FROM uploads;

USE transaction_validator;
TRUNCATE TABLE uploads;

