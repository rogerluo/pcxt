use course;
create table courseinfo(
	id int auto_increment not null primary key,
	courseid varchar(8) not null,
	coursename varchar(40) not null,
	coursetime int,
	coursecredit int not null,
	courseteacher varchar(10) not null,
	courseteacher2 varchar(10),
	courseheadcount int,
	coursetype varchar(15) not null,
	coursedate varchar(15) not null,
	coursefile varchar(256) not null
);