
Идея решения в том, чтобы развернуть все уникальные месяца (mm.yyyy), содержащиеся в таблице tbl_EVENTS в отдельные столбцы, сгруппировать все подписки/отписки пользователя в строку и на пересечении строки и столбца получить статус его подписки в текущем месяце: нет подписки, подписался, подписан, отписался, подписался и отписался в течении одного месяца.
В коде программы статусы указаны числами 0, 1, 2, 3, 4 соответственно.
Из такого датафрейма выходные данные можно получить серией простых группировок.
 
Для сохранения всех операций, относящихся к одному и тому же id, перед изменением вида таблицы через pivot, индексы строк и id дублируются в отдельный датафрейм, а после — присоединяются обратно.
На данном этапе имеются только статусы 0, 1, 3.
Затем строки группируются по столбцу 'Customer ID'. Появляется статус 4 (подписался и отписался в течении одного месяца).
Получившаяся таблица содержит столбец 'Customer ID' в качестве столбца индекса и N столбцов вида mm.yyyy.
 
Для получения статуса 2 (подписан) делается обход датафрейма. Все статусы 0, находящиеся между статусами 1 (подписался) и 3 (отписался), либо от начала строки и до статуса 3, либо от статуса 1 и до конца строки (при отсутствии информации о дате начала либо конца подписки во входных данных) заменяются на статусы 2 (подписан).
 
По формуле вычисляется общий customer_churn_rate без группировок, записывается в новый датафрейм.
 
Создается датафрейм из таблицы tbl_CUSTOMERS, объединяется с датафреймом, созданным из таблицы tbl_EVENTS.
 
По формуле вычисляются последовательно customer_churn_rate с группировками по стране, по типу подписки, записываются в новые датафреймы.
 
Три получившихся датафрейма записываются на листы «Без группировок», «По странам», «По типу подписки» в документ customer_churn_rate_file.xlsx
