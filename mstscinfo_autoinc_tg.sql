create or replace trigger mstscinfo_autoinc_tg
 before insert on mstscinfo for each row
 
begin
 select mstscinfo_autoinc_seq.nextval into :new.id from dual;
 end mstscinfo_autoinc_tg;
