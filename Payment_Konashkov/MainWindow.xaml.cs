using System;
using System.Windows.Forms;
using System.Drawing;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms.DataVisualization.Charting;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace Payment_Konashkov
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Entities _context = new Entities();

        public MainWindow()
        {
            InitializeComponent();
            this.Closing += MainWindow_Closing;

            // Инициализация графика
            var chart = new Chart();
            chart.ChartAreas.Add(new ChartArea("Main"));

            var currentSeries = new Series("Платежи") { IsValueShownAsLabel = true };
            chart.Series.Add(currentSeries);

            // Устанавливаем chart внутри WindowsFormsHost
            ChartPayments.Child = chart;

            CmbUser.ItemsSource = _context.User.ToList();
            CmbDiagram.ItemsSource = Enum.GetValues(typeof(SeriesChartType));
        }
        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // Показываем диалог с подтверждением выхода
            var result = System.Windows.MessageBox.Show("Вы уверены, что хотите выйти?", "Подтверждение", MessageBoxButton.YesNo, MessageBoxImage.Question);

            // Если пользователь выбрал "Нет", отменяем закрытие
            if (result != MessageBoxResult.Yes)
            {
                e.Cancel = true;  // Отменяем закрытие окна
            }
        }

        // Обновление диаграммы
        private void UpdateChart(object sender, SelectionChangedEventArgs e)
        {
            if (CmbUser.SelectedItem is User currentUser && CmbDiagram.SelectedItem is SeriesChartType currentType)
            {
                // Получаем ссылку на Chart внутри WindowsFormsHost
                var chart = ChartPayments.Child as Chart;
                var currentSeries = chart?.Series.FirstOrDefault();

                if (currentSeries != null)
                {
                    currentSeries.ChartType = currentType;
                    currentSeries.Points.Clear();

                    // Заполняем диаграмму данными
                    var categoriesList = _context.Category.ToList();

                    // Отладочная информация для категорий
                    Console.WriteLine($"Total categories: {categoriesList.Count}");

                    // Проверяем, если categoriesList пустой
                    if (categoriesList.Count == 0)
                    {
                        Console.WriteLine("No categories found.");
                        return;
                    }

                    // Проверяем количество категорий
                    int addedCategories = 0;

                    foreach (var category in categoriesList)
                    {
                        // Преобразуем результат в тип double? (nullable double)
                        var totalAmount = _context.Payment
                            .Where(p => p.UserID == currentUser.ID && p.CategoryID == category.ID)
                            .Sum(p => (double?)(p.Price * p.Num)); // Преобразуем в nullable double

                        // Если totalAmount null, заменяем на 0
                        double amount = totalAmount ?? 0;

                        // Отладочная информация для суммы и категории
                        Console.WriteLine($"Category: {category.Name}, Amount: {amount}");

                        // Добавляем точку на график
                        currentSeries.Points.AddXY(category.Name, amount);

                        addedCategories++;
                    }

                    // Проверка, добавились ли все категории
                    if (addedCategories < categoriesList.Count)
                    {
                        Console.WriteLine($"Warning: Only {addedCategories} categories were added, out of {categoriesList.Count}.");
                    }

                    // Настройка оси X для отображения всех категорий
                    var chartArea = chart.ChartAreas["Main"];
                    chartArea.AxisX.IsLabelAutoFit = false;  // Отключаем автофит
                    chartArea.AxisX.LabelStyle.Angle = 45;  // Устанавливаем угол наклона меток (например, 45 градусов)
                    chartArea.AxisX.Interval = 1;  // Устанавливаем интервал между метками равным 1, чтобы отображались все категории

                    // Используем другие параметры для меток оси X, без Font
                    chartArea.AxisX.LabelStyle.IsStaggered = true; // Ставит метки по диагонали (если метки слишком длинные)
                    chartArea.AxisX.LabelStyle.Format = "{0}"; // Устанавливает формат меток (если нужно, измените формат)
                }
            }
        }



        // Экспорт данных диаграммы в Excel
        private void ExportToExcelButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Получаем список всех пользователей, отсортированных по ФИО
                var allUsers = _context.User.ToList().OrderBy(u => u.FIO).ToList();

                // Создаем объект Excel
                var application = new Excel.Application
                {
                    SheetsInNewWorkbook = allUsers.Count(),
                    Visible = true // Чтобы приложение Excel было видно
                };

                // Добавляем новую книгу
                Excel.Workbook workbook = application.Workbooks.Add(Type.Missing);

                for (int i = 0; i < allUsers.Count(); i++)
                {
                    int startRowIndex = 1;
                    Excel.Worksheet worksheet = application.Worksheets.Item[i + 1];
                    worksheet.Name = allUsers[i].FIO; // Название листа - ФИО пользователя

                    // Заголовки столбцов
                    worksheet.Cells[1][startRowIndex] = "Дата платежа";
                    worksheet.Cells[2][startRowIndex] = "Название";
                    worksheet.Cells[3][startRowIndex] = "Стоимость";
                    worksheet.Cells[4][startRowIndex] = "Количество";
                    worksheet.Cells[5][startRowIndex] = "Сумма";

                    Excel.Range columnHeaderRange = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[5][1]];
                    columnHeaderRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    columnHeaderRange.Font.Bold = true;
                    startRowIndex++;

                    // Группируем платежи по категориям
                    var userCategories = allUsers[i].Payment
                        .OrderBy(u => u.Date)
                        .GroupBy(u => u.Category)
                        .OrderBy(u => u.Key.Name);

                    // Цикл по категориям платежей
                    foreach (var groupCategory in userCategories)
                    {
                        // Отображаем название категории
                        Excel.Range categoryHeaderRange = worksheet.Range[worksheet.Cells[1][startRowIndex], worksheet.Cells[5][startRowIndex]];
                        categoryHeaderRange.Merge();
                        categoryHeaderRange.Value = groupCategory.Key.Name;
                        categoryHeaderRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        categoryHeaderRange.Font.Italic = true;
                        startRowIndex++;

                        // Цикл по платежам в категории
                        foreach (var payment in groupCategory)
                        {
                            worksheet.Cells[1][startRowIndex] = payment.Date.ToString("dd.MM.yyyy");
                            worksheet.Cells[2][startRowIndex] = payment.Name;
                            worksheet.Cells[3][startRowIndex] = payment.Price;
                            (worksheet.Cells[3][startRowIndex] as Excel.Range).NumberFormat = "0.00";
                            worksheet.Cells[4][startRowIndex] = payment.Num;
                            worksheet.Cells[5][startRowIndex].Formula = $"=C{startRowIndex}*D{startRowIndex}";
                            (worksheet.Cells[5][startRowIndex] as Excel.Range).NumberFormat = "0.00";
                            startRowIndex++;
                        }

                        // Добавляем итог по категории
                        Excel.Range sumRange = worksheet.Range[worksheet.Cells[1][startRowIndex], worksheet.Cells[4][startRowIndex]];
                        sumRange.Merge();
                        sumRange.Value = "ИТОГО:";
                        sumRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                        worksheet.Cells[5][startRowIndex].Formula = $"=SUM(E{startRowIndex - groupCategory.Count()}:" +
                                                                  $"E{startRowIndex - 1})";
                        sumRange.Font.Bold = worksheet.Cells[5][startRowIndex].Font.Bold = true;
                        startRowIndex++;
                    }

                    // Добавляем границы для таблицы
                    Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[5][startRowIndex - 1]];
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle =
                        Excel.XlLineStyle.xlContinuous;

                    // Автоматическая подгонка ширины столбцов
                    worksheet.Columns.AutoFit();
                }

                // Добавление общего итога для всех пользователей на отдельном листе
                Excel.Worksheet totalWorksheet = application.Worksheets.Add();
                totalWorksheet.Name = "Общий итог";
                int rowIndex = 1;

                totalWorksheet.Cells[1][rowIndex] = "Общий итог:";
                totalWorksheet.Cells[2][rowIndex] = "Сумма всех платежей:";

                // Подсчитываем общую сумму по всем пользователям
                totalWorksheet.Cells[3][rowIndex].Formula = "=SUM(" + string.Join(",", allUsers.Select((user, index) => $"'{user.FIO}'!E2:E{index + 2}")) + ")";
                (totalWorksheet.Cells[3][rowIndex] as Excel.Range).NumberFormat = "0.00";

                // Форматируем строку общего итога красным цветом (RGB)
                Excel.Range totalRange = totalWorksheet.Range[totalWorksheet.Cells[1][rowIndex], totalWorksheet.Cells[3][rowIndex]];
                totalRange.Font.Color = 255; // Красный цвет

                // Отображаем Excel
                application.Visible = true;
            }
            catch (Exception ex)
            {
                // Обработка ошибок
                System.Windows.MessageBox.Show($"Произошла ошибка: {ex.Message}");
            }
        }


        // Экспорт данных диаграммы в Word
        private void ExportToWordButton_Click(object sender, EventArgs e)
        {
            var allUsers = _context.User.ToList();  // Получаем список пользователей
            var allCategories = _context.Category.ToList();  // Получаем список категорий

            // Создаем новый документ Word
            var application = new Word.Application();
            Word.Document document = application.Documents.Add();

            // Добавляем верхний колонтитул с текущей датой (Проверка наличия колонтитула)
            if (document.Sections.Count > 0)
            {
                var headerRange = document.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Text = $"Отчет о платежах на {DateTime.Now.ToString("dd.MM.yyyy")}";
                headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            }

            // Перебираем пользователей
            foreach (var user in allUsers)
            {
                // Проверяем, выбран ли пользователь в ComboBox
                if (CmbUser.SelectedItem != null && user == CmbUser.SelectedItem as User)
                {
                    // Добавляем абзац с именем пользователя
                    Word.Paragraph userParagraph = document.Paragraphs.Add();
                    Word.Range userRange = userParagraph.Range;
                    userRange.Text = user.FIO;
                    // Убедитесь, что стиль существует, например, "Заголовок1"
                    userParagraph.set_Style("Заголовок 1");
                    userRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    userRange.InsertParagraphAfter();

                    // Добавляем пустую строку
                    document.Paragraphs.Add();

                    // Добавляем таблицу с платежами
                    Word.Paragraph tableParagraph = document.Paragraphs.Add();
                    Word.Range tableRange = tableParagraph.Range;
                    Word.Table paymentsTable = document.Tables.Add(tableRange, allCategories.Count + 1, 2);

                    paymentsTable.Borders.InsideLineStyle = paymentsTable.Borders.OutsideLineStyle =
                        Word.WdLineStyle.wdLineStyleSingle;
                    paymentsTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                    // Названия колонок
                    Word.Range cellRange;

                    cellRange = paymentsTable.Cell(1, 1).Range;
                    cellRange.Text = "Категория";
                    cellRange = paymentsTable.Cell(1, 2).Range;
                    cellRange.Text = "Сумма расходов";

                    // Форматирование строк заголовка
                    paymentsTable.Rows[1].Range.Font.Name = "Times New Roman";
                    paymentsTable.Rows[1].Range.Font.Size = 14;
                    paymentsTable.Rows[1].Range.Bold = 1;
                    paymentsTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    // Заполнение таблицы данными для выбранного пользователя
                    for (int i = 0; i < allCategories.Count; i++)
                    {
                        var currentCategory = allCategories[i];

                        // Заполнение столбца с категориями
                        cellRange = paymentsTable.Cell(i + 2, 1).Range;
                        cellRange.Text = currentCategory.Name;
                        cellRange.Font.Name = "Times New Roman";
                        cellRange.Font.Size = 12;

                        // Заполнение столбца с суммой для выбранного пользователя
                        cellRange = paymentsTable.Cell(i + 2, 2).Range;
                        var totalAmount = user.Payment
                            .Where(p => p.Category == currentCategory)
                            .Sum(p => p.Num * p.Price);
                        cellRange.Text = $"{totalAmount:N2} руб.";  // Форматируем сумму
                        cellRange.Font.Name = "Times New Roman";
                        cellRange.Font.Size = 12;
                    }

                    // Добавляем пустую строку после таблицы
                    document.Paragraphs.Add();

                    // Добавляем информацию о самом дорогом платеже
                    var maxPayment = user.Payment.OrderByDescending(p => p.Price * p.Num).FirstOrDefault();
                    if (maxPayment != null)
                    {
                        Word.Paragraph maxPaymentParagraph = document.Paragraphs.Add();
                        Word.Range maxPaymentRange = maxPaymentParagraph.Range;
                        maxPaymentRange.Text = $"Самый дорогостоящий платеж - {maxPayment.Name} за {(maxPayment.Price * maxPayment.Num):N2} руб. от {maxPayment.Date:dd.MM.yyyy}";
                        maxPaymentParagraph.set_Style("Подзаголовок");
                        maxPaymentRange.Font.Color = Word.WdColor.wdColorDarkRed;
                        maxPaymentRange.InsertParagraphAfter();
                    }

                    // Добавляем информацию о самом дешевом платеже
                    var minPayment = user.Payment.OrderBy(p => p.Price * p.Num).FirstOrDefault();
                    if (minPayment != null)
                    {
                        Word.Paragraph minPaymentParagraph = document.Paragraphs.Add();
                        Word.Range minPaymentRange = minPaymentParagraph.Range;
                        minPaymentRange.Text = $"Самый дешевый платеж - {minPayment.Name} за {(minPayment.Price * minPayment.Num):N2} руб. от {minPayment.Date:dd.MM.yyyy}";
                        minPaymentParagraph.set_Style("Подзаголовок");
                        minPaymentRange.Font.Color = Word.WdColor.wdColorDarkGreen;
                        minPaymentRange.InsertParagraphAfter();
                    }

                    // Добавляем разрыв страницы, если это не последний пользователь
                    if (user != allUsers.LastOrDefault())
                    {
                        document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
                    }
                }
            }

            // Добавляем нижний колонтитул с номером страницы
            var footerRange = document.Sections[1].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
            footerRange.Fields.Add(footerRange, Word.WdFieldType.wdFieldPage);
            footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            // Открываем документ
            application.Visible = true;

            // Сохраняем документ в формате .docx и .pdf
            document.SaveAs2(@"C:\Users\user\Documents\Payments.docx");
            document.SaveAs2(@"C:\Users\user\Documents\Payments.pdf", Word.WdExportFormat.wdExportFormatPDF);
        }



    }
}
