Really Old Application With some really interesting Features in VB 6.
Feel free to use it for whatever you like.
Cheers.


to make connection succesfull in vista,
give all previlages to vb6.exe by rt. clicking -> compatibilty -> run as admin
ok
done
------------------------
just in case mscal.ocx is not installed
copy it to windows\system32
run cmd as admin
cd c:\windows\system32
regsvr32 mscal.ocx
this will do the work...
--------------------

    'TO ADD DOCTORS NAME IN THE DOCTOR NAME DROP DOWN LIST
    Set Doctor = Ado.Execute("select fname, lname from doctor d, person p where p.per_id = d.per_id")
    
    While Not Doctor.EOF
            'CONCATINATING FIRST AND LAST NAME OF THE DR.
            MsgBox "fname" & Doctor("fname")
            MsgBox "lname" & Doctor("lname")
            cmbDr.AddItem (Doctor(0) & " " & Doctor(1))
            Doctor.MoveNext
    Wend
-----------------------------------------

for datagrid to work
Firstly, you need to set a reference to Microsoft Active Data Objects Recordset 2.1 library (msador15.dll). Then add the DataGrid component into your VB6 toolbar and paste in the following code into the Form_Load event after adding a Grid (DataGrid1) to your form:
-------------------------------------


select per.fname, per.lname, pat.pat_id from person per,patient pat, appt ap where per.per_id = pat.per_id and pat.pat_id = 1 and ap.appt_time = '1000-1030' and ap.appt_date = '24-sep-09'

select pat_id, fname, lname from person per ,patient pat  where per.per_id = pat.per_id
select * from person per ,patient pat  where per.per_id = pat.per_id
select per.per_id, pat.pat_id, per.fname, per.lname, per.dob, per.sex, per.addr, per.ph, pat.height, pat.weight from person per ,patient pat  where per.per_id = pat.per_id
select disease from medical med, history his, patient pat where pat.pat_id = 1 and pat.pat_id = his.pat_id and med.med_id = his.med_id

insert all into person values (1, 'MAYANK', 'JAIN', '11-May-85', 'M', 'KATRAJ', 9371744602) into person values (2, 'AMIT', 'SINGHAL', '3-Feb-79', 'M', 'SATARA', 9171234591) into person values (3, 'ATUL', 'SHARMA', '4-Mar-75', 'M', 'SATARA',  9276892390) into person values (4, 'ANITA', 'KUKREJA', '5-DEC-60', 'F', 'AUNDH', 9716446021) into person values (5, 'SHALINI', 'KULKARNI', '6-Feb-78', 'F', 'FC ROAD', 9919127891) into person values (6, 'ANIL', 'VERMA', '11-MAR-82', 'M', 'KATRAJ', 9371742302) into person values (7, 'AMEEY', 'SHINOY', '3-Feb-79', 'F', 'SATARA', 9171123591) into person values (8, 'RAKESH', 'SHARMA', '4-Mar-72', 'M', 'KP',  9274592390) into person values (9, 'AMRUTA', 'DABADE', '5-DEC-66', 'F', 'JM ROAD', 9716946021) into person values (10, 'SHALU', 'REHMAN', '6-Feb-90', 'F', 'AUNDH', 9918997891) select * from dual;

insert all into patient values (1, 1, 150, 67) into patient values (2, 2, 170, 100) into patient values (3, 3, 187, 89) into patient values (4, 4, 156, 56) into patient values (5, 5, 178, 82) into patient values (6, 6, 183, 72) into patient values (7, 7, 192, 83) select * from dual;

insert all into doctor values (1, 8, 'HEART SPECIALIST' , 'PUNE UNIVERSITY') into doctor values (2, 9, 'GYNOCOLOGIST', 'OSMANIA UNIVERSITY') select * from dual;

insert all into staff values (1, 10,4000 , 'RECEPTIONIST') select * from dual;

insert all into test values (1, 'BLOOD', '12cc' , 100) into test values (2, 'CHEST XRAY', NULL , 100) into test values (3, 'URINE', '12gm' , 250) into test values (4, 'STOOL', '13cc' , 320) select * from dual;

insert all into medical values (1, 'OPERATION', 'HEART') into medical values (2, 'OPERATION', 'BRAIN') into medical values (3, 'OTHER', 'SMOKING') into medical values (4, 'OTHER', 'DRINKING') into medical values (5, 'CANCER', 'SKIN') into medical values (6, 'CANCER', 'LUNG') select * from dual;


create table person(per_id number(10),fname varchar(20),lname varchar(20),dob date,sex varchar(1),addr varchar(35),ph number(10),constraint pk_person primary key (per_id));

create table patient(pat_id number(10),per_id number(10),height number(3),weight number(3),constraint pk_patient primary key (pat_id),constraint fk_patient foreign key (per_id) references person(per_id));

create table doctor(doc_id number(10),per_id number(10),spec varchar(20),edu varchar(20),constraint pk_doctor primary key (doc_id),constraint fk_doctor foreign key (per_id) references person(per_id));

create table staff(staff_id number(10),per_id number(10),pay number(10),category varchar(20),constraint pk_staff primary key (staff_id),constraint fk_staff foreign key (per_id) references person(per_id));

create table test(test_id number(10),tname varchar(10),norm_result varchar(10),cost number(10),constraint pk_test primary key (test_id) );

create table test_result(test_id number(10),pat_id number(10),presc_id number(10),result varchar(10),constraint pk_test_result primary key (test_id, pat_id),constraint fk_test_result1 foreign key (test_id) references test(test_id),constraint fk_test_result2 foreign key (presc_id) references presc(presc_id),constraint fk_test_result3 foreign key (pat_id) references patient
(pat_id));

create table presc(presc_id number(10),doc_id number(10),pat_id number(10), med_recommended varchar(300),constraint pk_presc primary key (presc_id),constraint fk_presc1 foreign key (doc_id) references doctor(doc_id),constraint fk_presc2 foreign key (pat_id) references patient(pat_id));

create table appt (doc_id number(10),appt_date date, appt_time varchar(10), presc_id number(10),pat_id number(10) Not Null,constraint pk_appt primary key (doc_id, appt_date, appt_time),constraint fk_appt1 foreign key (doc_id) references doctor(doc_id), constraint fk_appt2 foreign key (presc_id) references presc(presc_id),constraint fk_appt3 foreign key (pat_id) references patient(pat_id));

create table medical (med_id number(10),category varchar(20),disease varchar(20),constraint pk_medical primary key (med_id));

create table history (med_id number(10),pat_id number(10),detail varchar(20),constraint pk_history primary key (med_id, pat_id),constraint fk_history1 foreign key (med_id) references medical(med_id),constraint fk_history2 foreign key (pat_id) references patient(pat_id));

---------------

drop table test_result
drop table test
drop table staff
drop table appt
drop table history

drop table id_table
drop table medical

drop table presc
drop table patient
drop table doctor
drop table person


to display all the tables in 10g 
select * from cat
