use course;
create table totalevaluate(
	id int auto_increment not null primary key,
	courseid varchar(8) not null,
	coursedate varchar(20) not null,
	itemid int not null,
	optionA varchar(12) not null,
	optionB varchar(12) not null,
	optionC varchar(12) not null,
	optionD varchar(12) not null
);