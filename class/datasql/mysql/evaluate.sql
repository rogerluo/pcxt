use course;
create table evaluate(
	id int auto_increment not null primary key,
	courseid varchar(8) not null,
	coursedate varchar(20) not null,
	stuid varchar(8) not null,
	stuname varchar(10) not null,
	book varchar(1) not null,
	courseraw varchar(1) not null,
	rawlang varchar(1) not null,
	lang varchar(1) not null,
	scale varchar(1) not null,
	spirit varchar(1) not null,
	info varchar(1) not null,
	teachstatus varchar(1) not null,
	teachtime varchar(1) not null,
	teachtime2 varchar(1),
	teachername2 varchar(10),
	stoptimes varchar(1),
	stopcourse varchar(20),
	retimes varchar(1),
	total varchar(1) not null,
	advice varchar(100)
	
);