use course
create table adminuser(
	id int IDENTITY(1,1) not null primary key,
	adminname varchar(10) not null,
	password varchar(100) not null,
	sysstate int not null
)
insert adminuser (adminname,password,sysstate) values
('dzyjs','ED800F10B45C7049FD0B3310E76CCCC4',0)