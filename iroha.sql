-- -*- mode: sql; sql-product: ms; coding: cp932-dos -*-
-- �����f�[�^���폜
drop table ta�����;
-- �f�[�^�����ݒ�
create table ta����� (�ԍ� counter, ������1 varchar, ������2 varchar, ���l integer);
-- ���f�[�^��ǉ�
insert into ta����� (������1, ������2, ���l) values ('����͂ɂقւ�', '����ʂ��', 2525);
insert into ta����� (������1, ������2, ���l) values ('�킩�悽�ꂻ' ,'�˂Ȃ��', 2828);
insert into ta����� (������1, ������2, ���l) values ('����̂������', '���ӂ�����', 28282525);
insert into ta����� (������1, ������2, ���l) values ('��������߂݂�', '��Ђ�������', 25252828);
-- �f�[�^�����o��
select * from ta�����;
select [������1] + '�@' + [������2] as ������ from ta�����;
