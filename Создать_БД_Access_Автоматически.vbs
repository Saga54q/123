' ==============================================================================
' Автоматический скрипт создания базы данных Microsoft Access (.accdb)
' Тема: "Поликлиника" (Вариант 14 - Сагадиев Амир)
' ==============================================================================

Option Explicit

Dim fso, currentDir, dbPath, accessApp, db

Set fso = CreateObject("Scripting.FileSystemObject")
currentDir = fso.GetParentFolderName(WScript.ScriptFullName)
dbPath = currentDir & "\Polyclinic_Database.accdb"

If fso.FileExists(dbPath) Then
    On Error Resume Next
    fso.DeleteFile dbPath, True
    If Err.Number <> 0 Then
        MsgBox "Пожалуйста, закройте открытую базу данных Polyclinic_Database.accdb перед повторным запуском!", vbExclamation, "Внимание"
        WScript.Quit
    End If
    On Error GoTo 0
End If

Set accessApp = CreateObject("Access.Application")
accessApp.NewCurrentDatabase dbPath
Set db = accessApp.CurrentDb

' -------------------------------------------------------------
' 1. СОЗДАНИЕ ТАБЛИЦ (Чистый стандартный синтаксис Access Jet/DAO)
' -------------------------------------------------------------

' 1) Таблица специальностей
db.Execute "CREATE TABLE Specialities (" & _
           "SpecialityID AUTOINCREMENT CONSTRAINT PK_Specialities PRIMARY KEY, " & _
           "SpecialityName TEXT(100) NOT NULL, " & _
           "Description TEXT(255));"
db.Execute "CREATE UNIQUE INDEX UQ_SpecialityName ON Specialities (SpecialityName);"

' 2) Таблица врачей
db.Execute "CREATE TABLE Doctors (" & _
           "DoctorID AUTOINCREMENT CONSTRAINT PK_Doctors PRIMARY KEY, " & _
           "FullName TEXT(100) NOT NULL, " & _
           "SpecialityID LONG NOT NULL, " & _
           "Cabinet TEXT(10) NOT NULL, " & _
           "Phone TEXT(20), " & _
           "Category TEXT(30), " & _
           "Status TEXT(20));"

' 3) Таблица пациентов
db.Execute "CREATE TABLE Patients (" & _
           "PatientID AUTOINCREMENT CONSTRAINT PK_Patients PRIMARY KEY, " & _
           "FullName TEXT(100) NOT NULL, " & _
           "BirthDate DATETIME NOT NULL, " & _
           "Gender TEXT(1) NOT NULL, " & _
           "Address TEXT(200), " & _
           "Phone TEXT(20) NOT NULL, " & _
           "OmsPolicy TEXT(16) NOT NULL);"
db.Execute "CREATE UNIQUE INDEX UQ_OmsPolicy ON Patients (OmsPolicy);"

' 4) Таблица медицинских услуг
db.Execute "CREATE TABLE MedicalServices (" & _
           "ServiceID AUTOINCREMENT CONSTRAINT PK_MedicalServices PRIMARY KEY, " & _
           "ServiceName TEXT(150) NOT NULL, " & _
           "SpecialityID LONG NOT NULL, " & _
           "Cost CURRENCY NOT NULL, " & _
           "DurationMinutes INTEGER);"
db.Execute "CREATE UNIQUE INDEX UQ_ServiceName ON MedicalServices (ServiceName);"

' 5) Таблица диагнозов (МКБ-10)
db.Execute "CREATE TABLE Diagnoses (" & _
           "DiagnosisID AUTOINCREMENT CONSTRAINT PK_Diagnoses PRIMARY KEY, " & _
           "MkbCode TEXT(10) NOT NULL, " & _
           "DiagnosisName TEXT(200) NOT NULL, " & _
           "Category TEXT(100));"
db.Execute "CREATE UNIQUE INDEX UQ_MkbCode ON Diagnoses (MkbCode);"

' 6) Таблица расписания приёма
db.Execute "CREATE TABLE DoctorSchedules (" & _
           "ScheduleID AUTOINCREMENT CONSTRAINT PK_DoctorSchedules PRIMARY KEY, " & _
           "DoctorID LONG NOT NULL, " & _
           "WorkDate DATETIME NOT NULL, " & _
           "StartTime DATETIME NOT NULL, " & _
           "EndTime DATETIME NOT NULL, " & _
           "Cabinet TEXT(10) NOT NULL);"

' 7) Таблица записей на приём (талонов)
db.Execute "CREATE TABLE Appointments (" & _
           "AppointmentID AUTOINCREMENT CONSTRAINT PK_Appointments PRIMARY KEY, " & _
           "TicketNumber TEXT(20) NOT NULL, " & _
           "DoctorID LONG NOT NULL, " & _
           "PatientID LONG NOT NULL, " & _
           "AppointmentDate DATETIME NOT NULL, " & _
           "AppointmentTime DATETIME NOT NULL, " & _
           "Status TEXT(20), " & _
           "Complaints MEMO);"
db.Execute "CREATE UNIQUE INDEX UQ_TicketNumber ON Appointments (TicketNumber);"

' 8) Таблица назначений
db.Execute "CREATE TABLE Prescriptions (" & _
           "PrescriptionID AUTOINCREMENT CONSTRAINT PK_Prescriptions PRIMARY KEY, " & _
           "AppointmentID LONG NOT NULL, " & _
           "DiagnosisID LONG NOT NULL, " & _
           "ServiceID LONG NOT NULL, " & _
           "PrescriptionText MEMO, " & _
           "Medications MEMO, " & _
           "Quantity INTEGER, " & _
           "Notes TEXT(255));"

' 9) Таблица учетных записей (ролевая безопасность)
db.Execute "CREATE TABLE AppUsers (" & _
           "UserID AUTOINCREMENT CONSTRAINT PK_AppUsers PRIMARY KEY, " & _
           "Username TEXT(50) NOT NULL, " & _
           "PasswordHash TEXT(64) NOT NULL, " & _
           "FullName TEXT(100) NOT NULL, " & _
           "Role TEXT(20) NOT NULL, " & _
           "IsActive YESNO);"
db.Execute "CREATE UNIQUE INDEX UQ_Username ON AppUsers (Username);"

' -------------------------------------------------------------
' 2. НАСТРОЙКА СВЯЗЕЙ И ССЫЛОЧНОЙ ЦЕЛОСТНОСТИ
' -------------------------------------------------------------
SubAddRelation db, "FK_Doctors_Speciality", "Specialities", "Doctors", "SpecialityID", "SpecialityID", 256
SubAddRelation db, "FK_Services_Speciality", "Specialities", "MedicalServices", "SpecialityID", "SpecialityID", 256
SubAddRelation db, "FK_Schedules_Doctor", "Doctors", "DoctorSchedules", "DoctorID", "DoctorID", 256 + 4096
SubAddRelation db, "FK_Appointments_Doctor", "Doctors", "Appointments", "DoctorID", "DoctorID", 256
SubAddRelation db, "FK_Appointments_Patient", "Patients", "Appointments", "PatientID", "PatientID", 256
SubAddRelation db, "FK_Prescriptions_App", "Appointments", "Prescriptions", "AppointmentID", "AppointmentID", 256 + 4096
SubAddRelation db, "FK_Prescriptions_Diag", "Diagnoses", "Prescriptions", "DiagnosisID", "DiagnosisID", 256
SubAddRelation db, "FK_Prescriptions_Serv", "MedicalServices", "Prescriptions", "ServiceID", "ServiceID", 256

' -------------------------------------------------------------
' 3. ЗАПОЛНЕНИЕ НАЧАЛЬНЫМИ ДАННЫМИ
' -------------------------------------------------------------
db.Execute "INSERT INTO Specialities (SpecialityName, Description) VALUES ('Терапевт', 'Первичный осмотр и общая терапия');"
db.Execute "INSERT INTO Specialities (SpecialityName, Description) VALUES ('Кардиолог', 'Заболевания сердечно-сосудистой системы');"
db.Execute "INSERT INTO Specialities (SpecialityName, Description) VALUES ('Невролог', 'Заболевания нервной системы');"
db.Execute "INSERT INTO Specialities (SpecialityName, Description) VALUES ('Офтальмолог', 'Коррекция зрения и патологии глаз');"
db.Execute "INSERT INTO Specialities (SpecialityName, Description) VALUES ('Хирург', 'Амбулаторная хирургия');"
db.Execute "INSERT INTO Specialities (SpecialityName, Description) VALUES ('Оториноларинголог', 'Заболевания уха, горла, носа');"

db.Execute "INSERT INTO Doctors (FullName, SpecialityID, Cabinet, Phone, Category, Status) VALUES ('Иванов Сергей Петрович', 1, '101', '+7 (917) 111-22-33', 'Высшая', 'Работает');"
db.Execute "INSERT INTO Doctors (FullName, SpecialityID, Cabinet, Phone, Category, Status) VALUES ('Смирнова Елена Васильевна', 1, '102', '+7 (917) 222-33-44', 'Первая', 'Работает');"
db.Execute "INSERT INTO Doctors (FullName, SpecialityID, Cabinet, Phone, Category, Status) VALUES ('Петров Алексей Михайлович', 2, '205', '+7 (917) 333-44-55', 'Высшая', 'Работает');"
db.Execute "INSERT INTO Doctors (FullName, SpecialityID, Cabinet, Phone, Category, Status) VALUES ('Кузнецова Ольга Дмитриевна', 3, '310', '+7 (917) 444-55-66', 'Высшая', 'Работает');"
db.Execute "INSERT INTO Doctors (FullName, SpecialityID, Cabinet, Phone, Category, Status) VALUES ('Ахметов Ильдар Равилевич', 4, '214', '+7 (917) 555-66-77', 'Первая', 'Работает');"
db.Execute "INSERT INTO Doctors (FullName, SpecialityID, Cabinet, Phone, Category, Status) VALUES ('Федоров Михаил Сергеевич', 5, '115', '+7 (917) 666-77-88', 'Высшая', 'Работает');"
db.Execute "INSERT INTO Doctors (FullName, SpecialityID, Cabinet, Phone, Category, Status) VALUES ('Хасанова Динара Фанилевна', 6, '208', '+7 (917) 777-88-99', 'Вторая', 'Работает');"

db.Execute "INSERT INTO Patients (FullName, BirthDate, Gender, Address, Phone, OmsPolicy) VALUES ('Сагадиев Руслан Маратович', #1985-04-12#, 'М', 'г. Уфа, пр. Октября, д. 45, кв. 12', '+7 (927) 100-20-30', '0254896321458790');"
db.Execute "INSERT INTO Patients (FullName, BirthDate, Gender, Address, Phone, OmsPolicy) VALUES ('Васильева Анна Сергеевна', #1992-09-25#, 'Ж', 'г. Уфа, ул. Ленина, д. 18, кв. 4', '+7 (927) 200-30-40', '0254896321458791');"
db.Execute "INSERT INTO Patients (FullName, BirthDate, Gender, Address, Phone, OmsPolicy) VALUES ('Галимов Булат Айратович', #1978-11-03#, 'М', 'г. Уфа, ул. Первомайская, д. 62, кв. 89', '+7 (927) 300-40-50', '0254896321458792');"
db.Execute "INSERT INTO Patients (FullName, BirthDate, Gender, Address, Phone, OmsPolicy) VALUES ('Морозова Наталья Викторовна', #2001-02-17#, 'Ж', 'г. Уфа, ул. Цюрупы, д. 91, кв. 34', '+7 (927) 400-50-60', '0254896321458793');"
db.Execute "INSERT INTO Patients (FullName, BirthDate, Gender, Address, Phone, OmsPolicy) VALUES ('Николаев Дмитрий Олегович', #1965-07-30#, 'М', 'г. Уфа, ул. Комсомольская, д. 15, кв. 77', '+7 (927) 500-60-70', '0254896321458794');"
db.Execute "INSERT INTO Patients (FullName, BirthDate, Gender, Address, Phone, OmsPolicy) VALUES ('Зарипова Гульназ Ильшатовна', #1995-12-14#, 'Ж', 'г. Уфа, ул. Менделеева, д. 122, кв. 15', '+7 (927) 600-70-80', '0254896321458795');"

db.Execute "INSERT INTO MedicalServices (ServiceName, SpecialityID, Cost, DurationMinutes) VALUES ('Первичный приём терапевта', 1, 900, 20);"
db.Execute "INSERT INTO MedicalServices (ServiceName, SpecialityID, Cost, DurationMinutes) VALUES ('Повторный приём терапевта', 1, 650, 15);"
db.Execute "INSERT INTO MedicalServices (ServiceName, SpecialityID, Cost, DurationMinutes) VALUES ('Электрокардиография (ЭКГ)', 2, 750, 15);"
db.Execute "INSERT INTO MedicalServices (ServiceName, SpecialityID, Cost, DurationMinutes) VALUES ('Консультация кардиолога расширенная', 2, 1400, 30);"
db.Execute "INSERT INTO MedicalServices (ServiceName, SpecialityID, Cost, DurationMinutes) VALUES ('Приём врача-невролога', 3, 1100, 20);"
db.Execute "INSERT INTO MedicalServices (ServiceName, SpecialityID, Cost, DurationMinutes) VALUES ('Проверка остроты зрения и подбор очков', 4, 800, 20);"
db.Execute "INSERT INTO MedicalServices (ServiceName, SpecialityID, Cost, DurationMinutes) VALUES ('Первичная хирургическая обработка раны', 5, 1250, 25);"
db.Execute "INSERT INTO MedicalServices (ServiceName, SpecialityID, Cost, DurationMinutes) VALUES ('Осмотр оториноларинголога и промывание лакун', 6, 950, 20);"

db.Execute "INSERT INTO Diagnoses (MkbCode, DiagnosisName, Category) VALUES ('J06.9', 'Острая респираторная вирусная инфекция (ОРВИ)', 'Болезни органов дыхания');"
db.Execute "INSERT INTO Diagnoses (MkbCode, DiagnosisName, Category) VALUES ('I10', 'Эссенциальная (первичная) гипертензия', 'Болезни системы кровообращения');"
db.Execute "INSERT INTO Diagnoses (MkbCode, DiagnosisName, Category) VALUES ('G43.0', 'Мигрень без ауры', 'Болезни нервной системы');"
db.Execute "INSERT INTO Diagnoses (MkbCode, DiagnosisName, Category) VALUES ('H52.1', 'Миопия (близорукость) слабой степени', 'Болезни глаза');"
db.Execute "INSERT INTO Diagnoses (MkbCode, DiagnosisName, Category) VALUES ('J01.0', 'Острый гайморит', 'Болезни органов дыхания');"
db.Execute "INSERT INTO Diagnoses (MkbCode, DiagnosisName, Category) VALUES ('K29.7', 'Гастрит неуточненный', 'Болезни органов пищеварения');"
db.Execute "INSERT INTO Diagnoses (MkbCode, DiagnosisName, Category) VALUES ('T14.0', 'Поверхностная травма тела', 'Травмы');"

db.Execute "INSERT INTO DoctorSchedules (DoctorID, WorkDate, StartTime, EndTime, Cabinet) VALUES (1, #2026-10-15#, #08:00:00#, #14:00:00#, '101');"
db.Execute "INSERT INTO DoctorSchedules (DoctorID, WorkDate, StartTime, EndTime, Cabinet) VALUES (2, #2026-10-15#, #14:00:00#, #20:00:00#, '102');"
db.Execute "INSERT INTO DoctorSchedules (DoctorID, WorkDate, StartTime, EndTime, Cabinet) VALUES (3, #2026-10-15#, #09:00:00#, #15:00:00#, '205');"
db.Execute "INSERT INTO DoctorSchedules (DoctorID, WorkDate, StartTime, EndTime, Cabinet) VALUES (4, #2026-10-15#, #08:30:00#, #14:30:00#, '310');"
db.Execute "INSERT INTO DoctorSchedules (DoctorID, WorkDate, StartTime, EndTime, Cabinet) VALUES (5, #2026-10-15#, #10:00:00#, #16:00:00#, '214');"
db.Execute "INSERT INTO DoctorSchedules (DoctorID, WorkDate, StartTime, EndTime, Cabinet) VALUES (6, #2026-10-15#, #08:00:00#, #13:00:00#, '115');"
db.Execute "INSERT INTO DoctorSchedules (DoctorID, WorkDate, StartTime, EndTime, Cabinet) VALUES (7, #2026-10-15#, #13:00:00#, #19:00:00#, '208');"

db.Execute "INSERT INTO Appointments (TicketNumber, DoctorID, PatientID, AppointmentDate, AppointmentTime, Status, Complaints) " & _
           "VALUES ('TAL-20261015-001', 1, 1, #2026-10-15#, #08:00:00#, 'Завершен', 'Кашель, температура 37.8, слабость');"
db.Execute "INSERT INTO Appointments (TicketNumber, DoctorID, PatientID, AppointmentDate, AppointmentTime, Status, Complaints) " & _
           "VALUES ('TAL-20261015-002', 1, 2, #2026-10-15#, #08:30:00#, 'Завершен', 'Головная боль, насморк в течение трех дней');"
db.Execute "INSERT INTO Appointments (TicketNumber, DoctorID, PatientID, AppointmentDate, AppointmentTime, Status, Complaints) " & _
           "VALUES ('TAL-20261015-003', 3, 3, #2026-10-15#, #09:00:00#, 'Завершен', 'Повышение давления до 160/100, одышка при ходьбе');"
db.Execute "INSERT INTO Appointments (TicketNumber, DoctorID, PatientID, AppointmentDate, AppointmentTime, Status, Complaints) " & _
           "VALUES ('TAL-20261015-004', 4, 4, #2026-10-15#, #09:30:00#, 'Завершен', 'Приступообразная пульсирующая боль в височной области');"
db.Execute "INSERT INTO Appointments (TicketNumber, DoctorID, PatientID, AppointmentDate, AppointmentTime, Status, Complaints) " & _
           "VALUES ('TAL-20261015-005', 5, 5, #2026-10-15#, #10:00:00#, 'Завершен', 'Ухудшение зрения вдаль, быстрая утомляемость глаз');"
db.Execute "INSERT INTO Appointments (TicketNumber, DoctorID, PatientID, AppointmentDate, AppointmentTime, Status, Complaints) " & _
           "VALUES ('TAL-20261015-006', 6, 1, #2026-10-15#, #10:30:00#, 'Завершен', 'Бытовая резаная рана предплечья');"
db.Execute "INSERT INTO Appointments (TicketNumber, DoctorID, PatientID, AppointmentDate, AppointmentTime, Status, Complaints) " & _
           "VALUES ('TAL-20261015-007', 7, 6, #2026-10-15#, #13:30:00#, 'Завершен', 'Заложенность носа, боль в области переносицы');"
db.Execute "INSERT INTO Appointments (TicketNumber, DoctorID, PatientID, AppointmentDate, AppointmentTime, Status, Complaints) " & _
           "VALUES ('TAL-20261015-008', 1, 4, #2026-10-15#, #11:00:00#, 'Запланирован', 'Плановый осмотр');"

db.Execute "INSERT INTO Prescriptions (AppointmentID, DiagnosisID, ServiceID, PrescriptionText, Medications, Quantity, Notes) " & _
           "VALUES (1, 1, 1, 'Постельный режим 4 дня, обильное питье', 'Парацетамол 500мг, Амоксициллин 500мг', 1, 'Больничный открыт');"
db.Execute "INSERT INTO Prescriptions (AppointmentID, DiagnosisID, ServiceID, PrescriptionText, Medications, Quantity, Notes) " & _
           "VALUES (2, 1, 1, 'Симптоматическая терапия', 'Спрей сосудосуживающий, Витамин С', 1, 'Повторный прием');"
db.Execute "INSERT INTO Prescriptions (AppointmentID, DiagnosisID, ServiceID, PrescriptionText, Medications, Quantity, Notes) " & _
           "VALUES (3, 2, 4, 'Контроль АД, диета с ограничением соли', 'Периндоприл 5мг, Индапамид', 1, 'Направлен на ЭКГ');"
db.Execute "INSERT INTO Prescriptions (AppointmentID, DiagnosisID, ServiceID, PrescriptionText, Medications, Quantity, Notes) " & _
           "VALUES (4, 3, 5, 'Соблюдение режима сна', 'Суматриптан 50мг', 1, 'Контроль через 14 дней');"
db.Execute "INSERT INTO Prescriptions (AppointmentID, DiagnosisID, ServiceID, PrescriptionText, Medications, Quantity, Notes) " & _
           "VALUES (5, 4, 6, 'Гимнастика для глаз', 'Капли Тауфон 4%, Лютеин', 1, 'Рецепт на очки');"
db.Execute "INSERT INTO Prescriptions (AppointmentID, DiagnosisID, ServiceID, PrescriptionText, Medications, Quantity, Notes) " & _
           "VALUES (6, 7, 7, 'Асептическая повязка', 'Хлоргексидин 0.05%, Левомеколь', 1, 'Перевязка');"
db.Execute "INSERT INTO Prescriptions (AppointmentID, DiagnosisID, ServiceID, PrescriptionText, Medications, Quantity, Notes) " & _
           "VALUES (7, 5, 8, 'Туалет носа, антибиотики', 'Аугментин 875/125мг', 1, 'Рентген пазух');"

db.Execute "INSERT INTO AppUsers (Username, PasswordHash, FullName, Role, IsActive) " & _
           "VALUES ('admin', 'admin123', 'Сагадиев Амир (Администратор)', 'Администратор', True);"
db.Execute "INSERT INTO AppUsers (Username, PasswordHash, FullName, Role, IsActive) " & _
           "VALUES ('user', 'user123', 'Регистратор поликлиники', 'Пользователь', True);"

' -------------------------------------------------------------
' 4. СОЗДАНИЕ ВСЕХ ТРЕБУЕМЫХ ЗАПРОСОВ (8 ЗАПРОСОВ)
' -------------------------------------------------------------
db.CreateQueryDef "qry_01_FreeDoctorSlots", _
    "PARAMETERS [pDate] DateTime, [pSpeciality] Text ( 100 ); " & _
    "SELECT Doctors.DoctorID, Doctors.FullName AS Врач, Specialities.SpecialityName AS Специальность, Doctors.Cabinet AS Кабинет, " & _
    "DoctorSchedules.WorkDate AS ДатаПриема, DoctorSchedules.StartTime AS НачалоСмены, DoctorSchedules.EndTime AS КонецСмены " & _
    "FROM (Specialities INNER JOIN Doctors ON Specialities.SpecialityID = Doctors.SpecialityID) " & _
    "INNER JOIN DoctorSchedules ON Doctors.DoctorID = DoctorSchedules.DoctorID " & _
    "WHERE (((Specialities.SpecialityName)=[pSpeciality]) AND ((DoctorSchedules.WorkDate)=[pDate]));"

db.CreateQueryDef "qry_02_PatientVisitHistory", _
    "PARAMETERS [pFullName] Text ( 100 ); " & _
    "SELECT Patients.FullName AS Пациент, Patients.BirthDate AS ДатаРождения, Patients.OmsPolicy AS ПолисОМС, " & _
    "Appointments.TicketNumber AS Талон, Appointments.AppointmentDate AS ДатаПриема, Appointments.AppointmentTime AS Время, " & _
    "Doctors.FullName AS ЛечащийВрач, Specialities.SpecialityName AS Специальность, " & _
    "Diagnoses.MkbCode AS КодМКБ, Diagnoses.DiagnosisName AS Диагноз, MedicalServices.ServiceName AS ОказаннаяУслуга, " & _
    "MedicalServices.Cost AS Стоимость, Prescriptions.Medications AS НазначенныеПрепараты " & _
    "FROM Specialities INNER JOIN (Doctors INNER JOIN (Patients INNER JOIN (Diagnoses INNER JOIN (MedicalServices " & _
    "INNER JOIN (Appointments INNER JOIN Prescriptions ON Appointments.AppointmentID = Prescriptions.AppointmentID) " & _
    "ON MedicalServices.ServiceID = Prescriptions.ServiceID) ON Diagnoses.DiagnosisID = Prescriptions.DiagnosisID) " & _
    "ON Patients.PatientID = Appointments.PatientID) ON Doctors.DoctorID = Appointments.DoctorID) " & _
    "ON Specialities.SpecialityID = Doctors.SpecialityID " & _
    "WHERE (((Patients.FullName) Like '*' & [pFullName] & '*')) " & _
    "ORDER BY Appointments.AppointmentDate DESC, Appointments.AppointmentTime DESC;"

db.CreateQueryDef "qry_03_AppointmentDetailedCalculated", _
    "SELECT Appointments.AppointmentID, Appointments.TicketNumber AS НомерТалона, Appointments.AppointmentDate AS ДатаПриема, " & _
    "Patients.FullName AS Пациент, Doctors.FullName AS Врач, Specialities.SpecialityName AS Специальность, " & _
    "MedicalServices.ServiceName AS НаименованиеУслуги, MedicalServices.Cost AS ЦенаУслуги, Prescriptions.Quantity AS Количество, " & _
    "[Cost]*[Quantity] AS [ИтоговаяСтоимость], " & _
    "IIf([Cost]*[Quantity]>1000, [Cost]*[Quantity]*0.9, [Cost]*[Quantity]) AS [СтоимостьСоСкидкой] " & _
    "FROM Specialities INNER JOIN (Doctors INNER JOIN (Patients INNER JOIN (MedicalServices INNER JOIN (Appointments " & _
    "INNER JOIN Prescriptions ON Appointments.AppointmentID = Prescriptions.AppointmentID) " & _
    "ON MedicalServices.ServiceID = Prescriptions.ServiceID) ON Patients.PatientID = Appointments.PatientID) " & _
    "ON Doctors.DoctorID = Appointments.DoctorID) ON Specialities.SpecialityID = Doctors.SpecialityID;"

db.CreateQueryDef "qry_04_DoctorWorkloadSummary", _
    "SELECT Specialities.SpecialityName AS Специальность, Doctors.FullName AS ФИО_Врача, Doctors.Cabinet AS Кабинет, " & _
    "Count(Appointments.AppointmentID) AS ВсегоПринятоПациентов, " & _
    "Sum(MedicalServices.Cost) AS ОбщаяСуммаОказанныхУслуг, " & _
    "Avg(MedicalServices.Cost) AS СредняяСтоимостьПриема " & _
    "FROM Specialities INNER JOIN (Doctors INNER JOIN (MedicalServices INNER JOIN (Appointments " & _
    "INNER JOIN Prescriptions ON Appointments.AppointmentID = Prescriptions.AppointmentID) " & _
    "ON MedicalServices.ServiceID = Prescriptions.ServiceID) ON Doctors.DoctorID = Appointments.DoctorID) " & _
    "ON Specialities.SpecialityID = Doctors.SpecialityID " & _
    "GROUP BY Specialities.SpecialityName, Doctors.FullName, Doctors.Cabinet " & _
    "ORDER BY Count(Appointments.AppointmentID) DESC;"

db.CreateQueryDef "qry_05_AllActiveDoctorsList", _
    "SELECT Doctors.DoctorID, Doctors.FullName AS [ФИО Врача], Specialities.SpecialityName AS Специальность, " & _
    "Doctors.Cabinet AS Кабинет, Doctors.Phone AS Телефон, Doctors.Category AS Категория " & _
    "FROM Specialities INNER JOIN Doctors ON Specialities.SpecialityID = Doctors.SpecialityID " & _
    "WHERE (((Doctors.Status)='Работает')) " & _
    "ORDER BY Specialities.SpecialityName, Doctors.FullName;"

db.CreateQueryDef "qry_06_AddNewPatientAppend", _
    "PARAMETERS [pFullName] Text ( 100 ), [pBirth] DateTime, [pGender] Text ( 1 ), [pAddress] Text ( 200 ), [pPhone] Text ( 20 ), [pOms] Text ( 16 ); " & _
    "INSERT INTO Patients ( FullName, BirthDate, Gender, Address, Phone, OmsPolicy ) " & _
    "VALUES ([pFullName], [pBirth], [pGender], [pAddress], [pPhone], [pOms]);"

db.CreateQueryDef "qry_07_UpdateServiceCostBatch", _
    "PARAMETERS [pPercent] IEEEDouble, [pSpecialityID] Long; " & _
    "UPDATE MedicalServices " & _
    "SET MedicalServices.Cost = Round([Cost]*(1+[pPercent]/100),2) " & _
    "WHERE (((MedicalServices.SpecialityID)=[pSpecialityID]));"

db.CreateQueryDef "qry_08_DeleteCancelledAppointments", _
    "PARAMETERS [pBeforeDate] DateTime; " & _
    "DELETE Appointments.*, Appointments.AppointmentDate, Appointments.Status " & _
    "FROM Appointments " & _
    "WHERE (((Appointments.AppointmentDate)<[pBeforeDate]) AND ((Appointments.Status)='Отменен'));"

accessApp.CloseCurrentDatabase
accessApp.Quit
Set accessApp = Nothing

MsgBox "База данных MS Access успешно создана:" & vbCrLf & dbPath & vbCrLf & vbCrLf & _
       "В базе созданы:" & vbCrLf & _
       "- 8 таблиц со связями и целостностью данных" & vbCrLf & _
       "- 8 готовых запросов" & vbCrLf & _
       "- Демонстрационные данные и таблица ролей", vbInformation, "Успешно создано!"

Sub SubAddRelation(oDb, sRelName, sPrimaryTbl, sForeignTbl, sPrimaryCol, sForeignCol, lAttr)
    Dim oRel, oFld
    Set oRel = oDb.CreateRelation(sRelName, sPrimaryTbl, sForeignTbl, lAttr)
    Set oFld = oRel.CreateField(sPrimaryCol)
    oFld.ForeignName = sForeignCol
    oRel.Fields.Append oFld
    oDb.Relations.Append oRel
End Sub
