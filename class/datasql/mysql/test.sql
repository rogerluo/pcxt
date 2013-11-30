use course
create table test(
	id int IDENTITY(1,1) not null primary key,
	username varchar(10) not null,
	password varchar(100) not null,
)