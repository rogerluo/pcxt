use course
create table coursemember(
	id int identity(1,1) not null primary key,
	stuid varchar(8) not null,
	stuname varchar(20) not null,
	stuobject varchar(20) not null,
	stucollege varchar(20) not null,
	courseid varchar(8) not null,
	coursename varchar(40),
	coursedate varchar(20) not null,
	isevaluate int not null
)