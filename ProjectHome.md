http://www.sqlite.org/cvstrac/wiki?p=ConverterTools

**Command line tool - Simple converter from MS Access fo SQLite**

The script unloads the structure and data of the Access database in a SQL file.

SQL includes:
  * create table ...
  * create index ...
  * alter table for foreign key constraints (not work in SQLite)
  * create trigger ... (instead of foreign keys, cascade too)
  * insert into ...