# IPS Viewer integration notes

Дата разведки: 2026-07-21.

## Цель

Нужно получить для ПО "Заказ заготовок ЦМО" read-only доступ к данным IPS и PDF-чертежам, которые сейчас просматриваются через установленное ClickOnce-приложение `IPS Viewer`.

## Найденное приложение

- Ярлык меню Пуск: `C:\Users\dpd\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\AO STP PZMC\IPS Viewer.appref-ms`.
- Тип установки: ClickOnce.
- Источник публикации из `.appref-ms`: `file://pz-files1/ПЗМЦ/18_ИТ/Каталог приложений/Внутренние приложения/ПО IPS Viewer/ips_viewer.application`.
- Прямое чтение сетевого каталога публикации из текущей сессии вернуло `Access is denied`; локальный ClickOnce-кэш читается.
- Локальная актуальная копия: `C:\Users\dpd\AppData\Local\Apps\2.0\R1NQW41M.J8O\ZDJ6CZ9J.VP8\ips_..tion_0000000000000000_0003.0000_f146d19a5688fcd6`.
- Основной exe: `ips_viewer.exe`, .NET Framework 4.7.2, WPF.

## Родственное приложение

В той же группе меню Пуск есть приложение `Производство.appref-ms`. Его локальная копия `pzmc.exe` содержит встроенный экран `ui/ips/ips_viewer.xaml`, а также разделы производства, техпроцессов и 1С OData. Это полезно как подтверждение, что IPS Viewer является частью общего внутреннего контура ПЗМЦ, но для нашей интеграции предпочтителен отдельный read-only адаптер, а не управление UI.

## Модель данных IPS по строкам сборки

Из `ips_viewer.exe` извлечены классы:

- `ips_viewer.Sql`: `Open`, `Close`, `LoadTypes`, `LoadObjects`, `LoadLinks`, `LoadAnalogs`, `LoadFiles`, `LoadIcons`, `LoadFile`.
- `ips_viewer.ObjectIPS`: объект IPS с полями обозначения, наименования, типа, состояния ЖЦ, серийного номера, ПС, заказчика, группы и файлов.
- `ips_viewer.LinkIPS`: связь родитель -> дочерний объект с количеством, единицей и версией.
- `ips_viewer.FileIPS`: файл IPS с `file_ips_id`, именем, размером, датой и `object_id`.

Основные таблицы/представления, которые использует IPS Viewer:

- `IMS_OBJECTS`
- `IMS_OBJECT_TYPES`
- `IMS_OBJECT_ATTRS`
- `IMS_RELATIONS`
- `IMS_RELATION_ATTRS`
- `IMS_STORAGE`
- `IMS_OBJTYPES_TREE`
- `IMS_LC_STEPS`
- `IMS_OBJECTS_VIEW`
- `IMS_POSSIBLE_VALUES`
- `IMV_A1536`
- `IMV_A5`

## Важные атрибуты

По запросам IPS Viewer:

- `IMS_OBJECT_ATTRS.F_ATTRIBUTE_ID = 9` - обозначение объекта.
- `IMS_OBJECT_ATTRS.F_ATTRIBUTE_ID = 10` - наименование объекта.
- `IMS_OBJECT_ATTRS.F_ATTRIBUTE_ID = 11` - примечание/описание.
- `IMS_OBJECT_ATTRS.F_ATTRIBUTE_ID = 1002` в `IMS_STORAGE` - файлы объекта, вероятно чертежи/PDF.
- `IMS_RELATION_ATTRS.F_ATTRIBUTE_ID = 1129` - количество в составе.
- `IMS_RELATION_ATTRS.F_ATTRIBUTE_ID = 1040` - версия детали в связи.
- `IMV_A5` + атрибут `13` - единица измерения количества.

Встречаются дополнительные атрибуты: `1146`, `10359`, `18226`, `18214`, `11101`, `18238`, `17776`, `18262`, `18283`, `18284`, `18285`, `18275`, `18302`. Их смысл нужно подтвердить живым запросом к IPS.

## Получение PDF

IPS Viewer не читает PDF как сетевые файлы напрямую. Он получает их из таблицы `IMS_STORAGE`.

Список файлов объекта:

```sql
SELECT
    s.F_FILE_ID,
    s.F_FILENAME,
    s.F_FILESIZE,
    s.F_FILEDATE,
    o.F_ID as F_OBJECT_ID
FROM IMS_STORAGE s
LEFT JOIN IMS_OBJECTS o ON (o.F_OBJECT_ID = s.F_OBJECTLINK_ID)
WHERE s.F_FILESIZE > 0
  AND s.F_ATTRIBUTE_ID = 1002
  AND s.F_AUTHOR > 0
  AND o.F_OBJECT_TYPE IN (...);
```

Тело файла:

```sql
SELECT TOP(1)
    F_FILESIZE,
    F_ARC_METHOD,
    F_ZIPSIZE,
    F_FILEBODY
FROM IMS_STORAGE
WHERE F_FILE_ID = @fileId;
```

Поле `F_FILEBODY` содержит бинарное тело файла. В сборке также есть признаки архивирования (`F_ARC_METHOD`, `F_ZIPSIZE`) и локальное сохранение в `.\files`. Поэтому адаптер должен уметь читать BLOB и, если потребуется, распаковывать тело по `F_ARC_METHOD`.

## Получение объектов и состава

Объекты IPS читаются из `IMS_OBJECTS` с присоединением атрибутов. Связи состава читаются из `IMS_RELATIONS` с типами связей `0`, `1`, `1071`, `1004`; аналоги - из связей типа `1060`.

Для нашей задачи это означает:

- поиск детали по IPS можно строить по `IMS_OBJECTS.F_ID`, `IMS_OBJECTS.F_OBJECT_ID`, `IMS_OBJECTS_VIEW.CAPTION`, обозначению из атрибута `9` и наименованию из атрибута `10`;
- состав детали/изделия можно читать через `IMS_RELATIONS`;
- PDF-чертежи можно привязывать к нашей детали по найденному `F_OBJECT_ID` / `F_ID` и списку `IMS_STORAGE`.

## Проверки подключения

- Попытка прочитать сетевую папку публикации ClickOnce вернула `Access is denied`.
- Прямой тест SQL Server `pz-sql1` через `System.Data.SqlClient` и Windows-аутентификацию завершился таймаутом pre-login handshake как в песочнице, так и вне песочницы. Это не подтверждает, что IPS живет на `pz-sql1`; сервер/БД/учетная запись пока неизвестны.
- В `ips_viewer.exe.config` строки подключения нет.
- В `ips_viewer.Sql.Open` строковые константы подключения не найдены; соединение, вероятно, формируется через внешний контекст, общий компонент, ClickOnce-активацию или другой скрытый источник.

## Рекомендуемый путь внедрения

1. Найти строку подключения к IPS:
   - через владельца/ИТ;
   - через документацию публикации `ips_viewer.application`;
   - через мониторинг уже запущенного IPS Viewer;
   - через деобфускацию/IL-разбор `Sql.Open` и вызовов `SqlConnection`.
2. Создать отдельный read-only helper в репозитории, по аналогии с `onec_integration`, но без записи в IPS:
   - `ips_integration/tools/ips_lookup.py` или C#-адаптер;
   - конфиг и учетные данные хранить вне репозитория;
   - команды: поиск IPS, список файлов, выгрузка PDF во временную папку/кэш.
3. Добавить в WPF-приложение сущности и UI-связь:
   - хранить внешний `IpsObjectId`/`IpsObjectCode` у `Part`;
   - показывать найденные PDF-чертежи в карточке детали/библиотеке;
   - кэшировать только метаданные и путь/хэш файла, не копировать массово все BLOB без команды пользователя.
4. Покрыть адаптер тестами на парсинг результатов и безопасный read-only режим. Живые запросы к IPS оставить ручными/интеграционными, чтобы не блокировать обычный `dotnet test`.

## Ограничения

На 2026-07-21 прямое получение данных IPS еще не реализовано: найдена структура и SQL-путь, но не найдено рабочее подключение к базе IPS. До подтверждения сервера/БД/прав нельзя встраивать живые запросы в основное приложение.
