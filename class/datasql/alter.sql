use course;

alter table courseinfo 
alter column courseid varchar(10)

alter table courseinfo 
add courseorder int;

alter table courseinfo
add constraint df_courseinfo_courseorder default 1 for courseorder;

alter table coursemember 
alter column stuid varchar(10);

alter table coursemember
alter column courseid varchar(10);

alter table coursemember
add courseorder int;

alter table coursemember
add constraint df_coursemember_courseorder default 1 for courseorder;

alter table evaluate
alter column stuid varchar(10);

alter table evaluate
alter column courseid varchar(10);

alter table evaluate
add courseorder int;

alter table evaluate
add constraint df_evaluate_courseorder default 1 for courseorder;

alter table totalevaluate
alter column courseid varchar(10);

alter table totalevaluate
add courseorder int;

alter table totalevaluate
add constraint df_totalevaluate_courseorder default 1 for courseorder;

GO