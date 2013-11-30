use course
create table evaluate_gra(
	id int identity(1,1) not null primary key,
	courseid varchar(8) not null,
	coursedate varchar(20) not null,
	stuid varchar(8) not null,
	stuname varchar(10) not null,
	coursecontent varchar(1) not null,
	coursedifficult varchar(1) not null,
	coursepoint varchar(1) not null,
	courselang varchar(1) not null,
	courseknowhelp varchar(1) not null,
	coursejobhelp varchar(1) not null,
	coursecontinue varchar(1) not null
)