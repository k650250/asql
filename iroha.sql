-- -*- mode: sql; sql-product: ms; coding: cp932-dos -*-
-- 既存データを削除
drop table taいろは;
-- データ書式設定
create table taいろは (番号 counter, 文字列1 varchar, 文字列2 varchar, 数値 integer);
-- 実データを追加
insert into taいろは (文字列1, 文字列2, 数値) values ('いろはにほへと', 'ちりぬるを', 2525);
insert into taいろは (文字列1, 文字列2, 数値) values ('わかよたれそ' ,'つねならむ', 2828);
insert into taいろは (文字列1, 文字列2, 数値) values ('うゐのおくやま', 'けふこえて', 28282525);
insert into taいろは (文字列1, 文字列2, 数値) values ('あさきゆめみし', 'ゑひもせすん', 25252828);
-- データを取り出す
select * from taいろは;
select [文字列1] + '　' + [文字列2] as 文字列 from taいろは;
