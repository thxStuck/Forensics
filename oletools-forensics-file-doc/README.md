<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>VBA Macro Extractor</title>
</head>
<body>
    <h1>VBA Macro Extractor</h1>

    <h2>Описание</h2>
    <p>Этот скрипт на Python предназначен для извлечения макросов из файлов Microsoft Office, используя библиотеку <code>oletools</code>. Он позволяет находить и отображать код VBA, содержащийся в документах, таких как Word и Excel. Это может быть полезно для анализа и аудита макросов, а также для обеспечения безопасности документов.</p>

    <h2>Установка</h2>
    <p>Для запуска скрипта необходимо установить Python и необходимые библиотеки. Убедитесь, что у вас установлен Python версии 3.6 или выше.</p>

    <h3>Шаги по установке:</h3>
    <ol>
        <li><strong>Скачайте и установите Python</strong>: <a href="https://www.python.org/downloads/">Скачать Python</a></li>
        <li><strong>Создайте виртуальное окружение (опционально)</strong>:
            <pre><code>python -m venv venv
source venv/bin/activate  # Для Linux/Mac
venv\Scripts\activate  # Для Windows</code></pre>
        </li>
        <li><strong>Установите необходимые библиотеки</strong>:
            <pre><code>pip install oletools</code></pre>
        </li>
    </ol>

    <h2>Использование</h2>
    <ol>
        <li>Сохраните скрипт в файл, например, <code>main.py</code>.</li>
        <li>Убедитесь, что у вас есть документ Microsoft Office (например, <code>cv_jones.doc_</code>), из которого вы хотите извлечь макросы.</li>
        <li>Измените переменную <code>file_name</code> в скрипте на имя вашего документа.</li>
        <li>Запустите скрипт:
            <pre><code>python main.py</code></pre>
        </li>
        <li>Скрипт выведет информацию о найденных макросах, включая имя файла, путь к макросу, имя модуля и сам код VBA.</li>
    </ol>

    <h2>Пример вывода</h2>
    <p>При успешном выполнении вы получите вывод, подобный следующему:</p>
    <pre><code>Макросы обнаружены.
Имя файла: cv_jones.doc_
Путь к макросу: Macros/VBA/ThisDocument
Имя модуля: ThisDocument.cls
Код VBA:
Sub Auto_Open()
    Anime
End Sub
...</code></pre>

    <h2>Важно</h2>
    <ul>
        <li>Используйте этот скрипт только для анализа документов, на которые у вас есть права. Извлечение и анализ макросов из чужих документов без разрешения может быть незаконным.</li>
        <li>Будьте осторожны с макросами, так как они могут содержать вредоносный код.</li>
    </ul>
</body>
</html>
