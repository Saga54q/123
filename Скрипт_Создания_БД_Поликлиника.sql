-- =============================================================================
-- СКРИПТ СОЗДАНИЯ СТРУКТУРЫ БАЗЫ ДАННЫХ MS ACCESS (РУССКИЕ ТАБЛИЦЫ И ПОЛЯ)
-- ПРЕДМЕТНАЯ ОБЛАСТЬ: ПОЛИКЛИНИКА (ВАРИАНТ 14)
-- СТУДЕНТ: САГАДИЕВ АМИР
-- =============================================================================

CREATE TABLE [Специальности] (
    [КодСпециальности] AUTOINCREMENT CONSTRAINT PK_Специальности PRIMARY KEY,
    [НаименованиеСпециальности] TEXT(100) NOT NULL,
    [Описание] TEXT(255)
);
CREATE UNIQUE INDEX UQ_Специальность ON [Специальности] ([НаименованиеСпециальности]);

CREATE TABLE [Врачи] (
    [КодВрача] AUTOINCREMENT CONSTRAINT PK_Врачи PRIMARY KEY,
    [ФИО_Врача] TEXT(100) NOT NULL,
    [КодСпециальности] LONG NOT NULL,
    [Кабинет] TEXT(10) NOT NULL,
    [Телефон] TEXT(20),
    [Категория] TEXT(30),
    [Статус] TEXT(20)
);

CREATE TABLE [Пациенты] (
    [КодПациента] AUTOINCREMENT CONSTRAINT PK_Пациенты PRIMARY KEY,
    [ФИО_Пациента] TEXT(100) NOT NULL,
    [ДатаРождения] DATETIME NOT NULL,
    [Пол] TEXT(1) NOT NULL,
    [Адрес] TEXT(200),
    [Телефон] TEXT(20) NOT NULL,
    [ПолисОМС] TEXT(16) NOT NULL
);
CREATE UNIQUE INDEX UQ_ПолисОМС ON [Пациенты] ([ПолисОМС]);

CREATE TABLE [МедицинскиеУслуги] (
    [КодУслуги] AUTOINCREMENT CONSTRAINT PK_МедицинскиеУслуги PRIMARY KEY,
    [НаименованиеУслуги] TEXT(150) NOT NULL,
    [КодСпециальности] LONG NOT NULL,
    [Стоимость] CURRENCY NOT NULL,
    [ДлительностьМинут] INTEGER
);
CREATE UNIQUE INDEX UQ_Услуга ON [МедицинскиеУслуги] ([НаименованиеУслуги]);

CREATE TABLE [Диагнозы] (
    [КодДиагноза] AUTOINCREMENT CONSTRAINT PK_Диагнозы PRIMARY KEY,
    [КодМКБ] TEXT(10) NOT NULL,
    [НаименованиеДиагноза] TEXT(200) NOT NULL,
    [Категория] TEXT(100)
);
CREATE UNIQUE INDEX UQ_КодМКБ ON [Диагнозы] ([КодМКБ]);

CREATE TABLE [РасписаниеПриема] (
    [КодРасписания] AUTOINCREMENT CONSTRAINT PK_РасписаниеПриема PRIMARY KEY,
    [КодВрача] LONG NOT NULL,
    [ДатаСмены] DATETIME NOT NULL,
    [ВремяНачала] DATETIME NOT NULL,
    [ВремяОкончания] DATETIME NOT NULL,
    [Кабинет] TEXT(10) NOT NULL
);

CREATE TABLE [ЗаписиНаПрием] (
    [КодЗаписи] AUTOINCREMENT CONSTRAINT PK_ЗаписиНаПрием PRIMARY KEY,
    [НомерТалона] TEXT(20) NOT NULL,
    [КодВрача] LONG NOT NULL,
    [КодПациента] LONG NOT NULL,
    [ДатаПриема] DATETIME NOT NULL,
    [ВремяПриема] DATETIME NOT NULL,
    [Статус] TEXT(20),
    [Жалобы] MEMO
);
CREATE UNIQUE INDEX UQ_НомерТалона ON [ЗаписиНаПрием] ([НомерТалона]);

CREATE TABLE [Назначения] (
    [КодНазначения] AUTOINCREMENT CONSTRAINT PK_Назначения PRIMARY KEY,
    [КодЗаписи] LONG NOT NULL,
    [КодДиагноза] LONG NOT NULL,
    [КодУслуги] LONG NOT NULL,
    [Рекомендации] MEMO,
    [Медикаменты] MEMO,
    [Количество] INTEGER,
    [Примечания] TEXT(255)
);

CREATE TABLE [Пользователи] (
    [КодПользователя] AUTOINCREMENT CONSTRAINT PK_Пользователи PRIMARY KEY,
    [Логин] TEXT(50) NOT NULL,
    [Пароль] TEXT(64) NOT NULL,
    [ФИО] TEXT(100) NOT NULL,
    [Роль] TEXT(20) NOT NULL,
    [Активен] YESNO
);
CREATE UNIQUE INDEX UQ_Логин ON [Пользователи] ([Логин]);
