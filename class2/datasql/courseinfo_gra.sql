use course
create table courseinfo_gra(
	id int identity(1,1) not null primary key,
	courseid varchar(8) not null,
	coursename varchar(40) not null,
	coursetime int not null,
	coursecredit int not null,
	courseteacher varchar(10) not null,
	courseheadcount int,
	coursetype varchar(15) not null,
	coursetype2 varchar(15) not null,
	coursedate varchar(15) not null
)