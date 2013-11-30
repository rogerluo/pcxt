use course
create table evaluate2_gra(
	id int identity(1,1) not null primary key,
	stuid varchar(8) not null,
	stuname varchar(10) not null,
	coursestudy text,
	coursejob text
)