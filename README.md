# Проект-помощник в работе 1С УТ (SQL MS Access)
Проект делался, как инструмент наведения порядка в базе "1С Управление торговлей", но может быстро быть адаптирован к любой другой версии 1С.
Почему не сразу в 1С ? Потому, что применять MS Access - удобнее и быстрее, и самое важное, безопасно для единства базы данных.
# Первыми задачами были:
- поле "Рабочее наименование" должно начинаться с существительного (правило фирмы);
- удаление дублей карточек.

Выявленные карточки редактируются в MS Access. Обновляются наименования, проставляется признак удаления в дублях. Таблица с отредактированными экспортируется в Excel. Затем импортируется в 1С.
# Импорт в 1С большого количества новой номенклатуры (бывало более тысячи за день, но можно и больше):
- часть товара уже может быть в 1С давно, но MS Access поможет избежать дублей.
# Сопровождение товаров нормативной документацией:
Медицинские товары сопровождаются регистрационными удостоверениями Здравнадзора (РУ), это влияет на ставку НДС. Лабораторное оборудование декларациями и сертификатами соответствия. Измерительное оборудование свидетельствами - сайт "АРШИН".
- программа отслеживает изменения в тех РУ, которые надо контролировать;
- сертификаты и свидетельства имеют сроки обращения, наступит день и их надо будет заменить или изъять из карточек.
# Другие обработки:
- аналогичная работа с базой 1С казахстанского филиала;
- по габаритам товаров и товарным местам;
- отслеживание отметки ГТД в карточках импорных товаров;
- ... прослеживаемый товар;
- оперативные аналитические данные по запросам руководства и коллег.
