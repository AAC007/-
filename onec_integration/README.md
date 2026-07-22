# Интеграция с 1С:УТ 10.3

Проект готовит безопасный контур интеграции с 1С:Предприятие 8.3 / УТ 10.3.47.1:

- обследование метаданных без создания документов;
- чтение складов, номенклатуры и остатков через штатный COM-коннектор 1С;
- подготовка документов комплектования и перемещения в режиме `DRY_RUN`;
- запись непроведенных документов только при явном `ALLOW_DOCUMENT_WRITE=true`;
- проведение только отдельной командой и только при `ALLOW_DOCUMENT_POSTING=true`.

## Архитектура

Предпочтительный промышленный вариант - HTTP-сервис в расширении 1С, потому что он централизует права, журналирование и публикацию API внутри платформы. До публикации HTTP-сервиса этот каталог дает COM-клиент для обследования базы и чтения остатков без изменения типовой конфигурации.

Запись напрямую в SQL не используется. Для остатков УТ 10.3 уже подтвержден основной регистр `ТоварыНаСкладах`; контрольные регистры для сверки - `ПартииТоваровНаСкладах` и `ТоварыОрганизаций`.

## Оперативный запрос остатка

Для точечных вопросов пользователя об остатках используй быстрый read-only путь:

```powershell
.\onec_integration\start_stock_lookup_server.cmd
.\onec_integration\query_stock_server.cmd -Item "01-P002-03.002-T63.150.SS" -Warehouse "44"
```

Helper-сервер использует `C:\Users\dpd\Documents\Codex\1C_Diagnostics\.env`, не выводит пароль, при проблемах с `pz-sql1.pzmc.org:1541` подключается через `pz-sql1:1541`.

По умолчанию читается только основной регистр `ТоварыНаСкладах`. Для сверки добавьте `-Control`. Для списка до 30 строк используйте один пакетный запуск:

```powershell
.\onec_integration\query_stock_server.cmd -Item "11445" "УТ000012706" -Warehouse "44"
```

Подробности: `docs/ONEC_STOCK_FAST_PATH.md`.

## Быстрый старт

1. Скопируйте `.env.example` в `.env` и заполните только `ONEC_PASSWORD`.
2. Скопируйте `onec_integration/config.example.json` в `onec_integration/config.json`.
3. Проверьте окружение:

```powershell
python -m onec_integration.cli diagnose-environment
```

4. Снимите отчет по метаданным:

```powershell
python -m onec_integration.cli inspect-metadata --out onec_integration/reports/metadata.json
```

5. Проверьте, что в `config.json` сохранены реальные имена регистра остатков и полей.
6. Прочитайте остатки:

```powershell
python -m onec_integration.cli stocks --positive-only --out onec_integration/reports/stocks.json
```

## Команды

```powershell
python -m onec_integration.cli diagnose-environment
python -m onec_integration.cli inspect-metadata --out onec_integration/reports/metadata.json
python -m onec_integration.cli warehouses --out onec_integration/reports/warehouses.json
python -m onec_integration.cli nomenclature --name "круг" --out onec_integration/reports/nomenclature.json
python -m onec_integration.cli stocks --positive-only --out onec_integration/reports/stocks.json
python -m unittest discover onec_integration/tests
```

## Риски и ограничения

- COM-подключение подтверждено через `C:\Users\dpd\Documents\Codex\1C_Diagnostics`; если FQDN не разрешается, используется `pz-sql1:1541`.
- COM-коннектор должен быть зарегистрирован в той же разрядности, что и Python.
- Документы в рабочей базе нельзя создавать тестами. Для проверки записи используйте отдельные контрольные операции с уникальным `external_id` и убирайте созданные непроведенные документы вручную через 1С.
- Без заполненного `config.json` модуль не строит запрос остатков, чтобы не угадывать регистры УТ.

## Готовность к внедрению в ПО

План замены Excel-импорта остатков прямой синхронизацией 1С подготовлен в `docs/STOCK_SYNC_IMPLEMENTATION_PLAN.md`. До отдельной команды пользователя WPF-приложение не меняется и продолжает использовать существующий импорт Excel.

## Что требуется от администратора 1С

- Подтвердить доступность сервера/базы и DNS-имени с ПК пользователя.
- Проверить регистрацию COM-коннектора `V83.COMConnector`.
- Выдать интеграционному пользователю минимальные права на чтение складов, номенклатуры, регистров остатков.
- Для этапа записи отдельно выдать права на запись непроведенных документов комплектования/перемещения.
- При выборе HTTP-варианта опубликовать расширение или внешнюю обработку на тестовом веб-сервере 1С.

