using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Export;

namespace ConsoleTestPptx;

public class PresentationMaker
{
	private readonly PresentationScope _presentationScope;

	public PresentationMaker(PresentationScope presentationScope)
	{
		_presentationScope = presentationScope;
	}

	public async Task BuildPresentation()
	{
		FormatFirstSlide();
		AddTableSlide(4);
		AddTableSlide(20);
		AddColumnChartSlide(2, 6);
		AddColumnChartSlide(8, 6);
		AddColumnChartSlide(2, 30);
		AddColumnChartSlide(8, 30);
		AddPieChartSlide(new Dictionary<string, int> 
		{ 
			{ "Продукты", 35 }, 
			{ "Услуги", 25 }, 
			{ "Консалтинг", 20 }, 
			{ "Поддержка", 15 }, 
			{ "Другое", 5 } 
		});
		AddPieChartSlide(new Dictionary<string, int> 
		{ 
			{ "Продукты", 35000 }, 
			{ "Услуги", 2500 }, 
			{ "Консалтинг", 200 }, 
			{ "Поддержка", 15 }, 
			{ "Другое", 5 } 
		});
		AddPieChartSlide(new Dictionary<string, int> 
		{ 
			{ "Продукты", 35000 }, 
			{ "Услуги", 2500 }, 
			{ "Консалтинг", 200 }, 
			{ "Поддержка", 15 }, 
			{ "Другое", 5 }, 
			{ "Продукты2", 35000 }, 
			{ "Услуги2", 2500 }, 
			{ "Консалтинг2", 200 }, 
			{ "Поддержка2", 15 }, 
			{ "Другое2", 5 } , 
			{ "Продукты3", 35000 }, 
			{ "Услуги3", 2500 }, 
			{ "Консалтинг3", 200 }, 
			{ "Поддержка3", 15 }, 
			{ "Другое3", 5 } 
		});
		AddLineChartSlide(2, 8);
		AddLineChartSlide(30, 8);
		AddLineChartSlide(2, 30);
		AddLineChartSlide(30, 40);
		AddAreaChartSlide(3, 5);
		AddAreaChartSlide(30, 5);
		AddAreaChartSlide(3, 30);
		AddAreaChartSlide(30, 40);
		AddScatterChartSlide();
		AddCombinedChartSlide();
		var filename = await CreateFilename();
		await SavePresentation(filename);
	}

	private void FormatFirstSlide()
	{
		var slide = _presentationScope.Presentation.Slides.Count > 0
			? _presentationScope.Presentation.Slides[0]
			: _presentationScope.Presentation.Slides.AddEmptySlide(_presentationScope.Presentation.LayoutSlides[0]);

		var slideSize = _presentationScope.Presentation.SlideSize.Size;
		var textWidth = 600f;
		var textHeight = 100f;
		var x = ((float)slideSize.Width - textWidth) / 2;
		var y = ((float)slideSize.Height - textHeight) / 2;
		
		var textBox = slide.Shapes.AddAutoShape(ShapeType.Rectangle, x, y, textWidth, textHeight);
		textBox.FillFormat.FillType = FillType.NoFill;
		textBox.LineFormat.FillFormat.FillType = FillType.NoFill;
		textBox.TextFrame.Text = "Тестовый тайтл";
		textBox.TextFrame.Paragraphs[0].ParagraphFormat.Alignment = TextAlignment.Center;
		var portion = textBox.TextFrame.Paragraphs[0].Portions[0];
		portion.PortionFormat.FontHeight = 48;
		portion.PortionFormat.LatinFont = new FontData("Calibri");
		portion.PortionFormat.FillFormat.FillType = FillType.Solid;
		portion.PortionFormat.FillFormat.SolidFillColor.Color = System.Drawing.Color.Black;

		var secondTextWidth = 400f;
		var secondTextHeight = 70f;
		var secondX = ((float)slideSize.Width - secondTextWidth) / 2;
		var secondY = y + textHeight + 50f;
		
		var secondTextBox = slide.Shapes.AddAutoShape(ShapeType.Rectangle, secondX, secondY, secondTextWidth, secondTextHeight);
		secondTextBox.FillFormat.FillType = FillType.NoFill;
		secondTextBox.LineFormat.FillFormat.FillType = FillType.NoFill;
		secondTextBox.Rotation = -30f;
		secondTextBox.TextFrame.Text = "Второй текст";
		secondTextBox.TextFrame.Paragraphs[0].ParagraphFormat.Alignment = TextAlignment.Center;
		var secondPortion = secondTextBox.TextFrame.Paragraphs[0].Portions[0];
		secondPortion.PortionFormat.FontHeight = 36;
		secondPortion.PortionFormat.LatinFont = new FontData("Times New Roman");
		secondPortion.PortionFormat.FillFormat.FillType = FillType.Solid;
		secondPortion.PortionFormat.FillFormat.SolidFillColor.Color = System.Drawing.Color.Purple;

		var thirdTextWidth = 800f;
		var thirdTextHeight = 100f;
		var thirdX = (float)slideSize.Width - thirdTextWidth;
		var thirdY = 0f;
		
		var thirdTextBox = slide.Shapes.AddAutoShape(ShapeType.Rectangle, thirdX, thirdY, thirdTextWidth, thirdTextHeight);
		thirdTextBox.FillFormat.FillType = FillType.NoFill;
		thirdTextBox.LineFormat.FillFormat.FillType = FillType.NoFill;
		thirdTextBox.Rotation = 45f;
		thirdTextBox.TextFrame.Text = "Очень длинный текст, который должен выйти за рамки слайда и быть видимым даже при повороте на 45 градусов";
		thirdTextBox.TextFrame.Paragraphs[0].ParagraphFormat.Alignment = TextAlignment.Left;
		var thirdPortion = thirdTextBox.TextFrame.Paragraphs[0].Portions[0];
		thirdPortion.PortionFormat.FontHeight = 36;
		thirdPortion.PortionFormat.LatinFont = new FontData("Arial");
		thirdPortion.PortionFormat.FillFormat.FillType = FillType.Solid;
		thirdPortion.PortionFormat.FillFormat.SolidFillColor.Color = System.Drawing.Color.Black;
	}

	private void AddTableSlide(int rowCount)
	{
		var slide = _presentationScope.Presentation.Slides.AddEmptySlide(_presentationScope.Presentation.LayoutSlides[0]);
		var slideSize = _presentationScope.Presentation.SlideSize.Size;
		
		IAutoShape? titleShape = null;
		foreach (IShape shape in slide.Shapes)
		{
			if (shape is IAutoShape autoShape && autoShape.TextFrame != null)
			{
				if (shape.Placeholder != null && 
				    (shape.Placeholder.Type == PlaceholderType.Title || shape.Placeholder.Type == PlaceholderType.CenteredTitle))
				{
					titleShape = autoShape;
					break;
				}
			}
		}
		
		if (titleShape == null)
		{
			foreach (IShape shape in slide.Shapes)
			{
				if (shape is IAutoShape autoShape && autoShape.TextFrame != null && shape.Y < 200)
				{
					titleShape = autoShape;
					break;
				}
			}
		}
		
		if (titleShape != null)
		{
			titleShape.TextFrame.Text = "название таблицы";
		}
		
		var tableMarginCm = 2.0;
		var tableMarginPoints = tableMarginCm * 28.35;
		var tableWidth = (float)slideSize.Width - (float)tableMarginPoints;
		var textColumnWidth = tableWidth * 0.3;
		var dateColumnWidth = tableWidth * 0.3;
		var numberColumnWidth = (tableWidth - textColumnWidth - dateColumnWidth) / 3;
		
		var tableX = ((float)slideSize.Width - tableWidth) / 2;
		var tableY = 150f;
		
		double[] columnWidths = { textColumnWidth, dateColumnWidth, numberColumnWidth, numberColumnWidth, numberColumnWidth };
		var rowHeights = new double[rowCount + 1];
		for (int i = 0; i < rowHeights.Length; i++)
		{
			rowHeights[i] = 50;
		}
		
		var table = slide.Shapes.AddTable(tableX, tableY, columnWidths, rowHeights);
		
		var headerRow = table.Rows[0];
		headerRow[0].TextFrame.Text = "Текст";
		headerRow[1].TextFrame.Text = "Дата и время";
		headerRow[2].TextFrame.Text = "Число 1";
		headerRow[3].TextFrame.Text = "Число 2";
		headerRow[4].TextFrame.Text = "Число 3";
		
		for (int col = 0; col < 5; col++)
		{
			var cell = headerRow[col];
			cell.CellFormat.FillFormat.FillType = FillType.Solid;
			cell.CellFormat.FillFormat.SolidFillColor.Color = System.Drawing.Color.LightGray;
			cell.TextFrame.Paragraphs[0].ParagraphFormat.Alignment = TextAlignment.Center;
			var portion = cell.TextFrame.Paragraphs[0].Portions[0];
			portion.PortionFormat.FontHeight = 14;
			portion.PortionFormat.FontBold = NullableBool.True;
		}
		
		for (int row = 1; row <= rowCount; row++)
		{
			var dataRow = table.Rows[row];
			dataRow[0].TextFrame.Text = $"Текст {row}";
			dataRow[1].TextFrame.Text = DateTime.Now.AddDays(row).ToString("dd.MM.yyyy HH:mm");
			dataRow[2].TextFrame.Text = (row * 10).ToString();
			dataRow[3].TextFrame.Text = (row * 20).ToString();
			dataRow[4].TextFrame.Text = (row * 30).ToString();
			
			var textCell = dataRow[0];
			textCell.TextFrame.Paragraphs[0].ParagraphFormat.Alignment = TextAlignment.Left;
			textCell.TextAnchorType = TextAnchorType.Top;
			var textPortion = textCell.TextFrame.Paragraphs[0].Portions[0];
			textPortion.PortionFormat.FontHeight = 12;
			
			var dateCell = dataRow[1];
			dateCell.TextFrame.Paragraphs[0].ParagraphFormat.Alignment = TextAlignment.Center;
			dateCell.TextAnchorType = TextAnchorType.Center;
			var datePortion = dateCell.TextFrame.Paragraphs[0].Portions[0];
			datePortion.PortionFormat.FontHeight = 12;
			datePortion.PortionFormat.FillFormat.FillType = FillType.Solid;
			datePortion.PortionFormat.FillFormat.SolidFillColor.Color = System.Drawing.Color.Red;
			
			var numberCell1 = dataRow[2];
			numberCell1.TextFrame.Paragraphs[0].ParagraphFormat.Alignment = TextAlignment.Right;
			numberCell1.TextAnchorType = TextAnchorType.Bottom;
			var numberPortion1 = numberCell1.TextFrame.Paragraphs[0].Portions[0];
			numberPortion1.PortionFormat.FontHeight = 12;
			numberPortion1.PortionFormat.LatinFont = new FontData("Arial");
			
			var numberCell2 = dataRow[3];
			numberCell2.TextFrame.Paragraphs[0].ParagraphFormat.Alignment = TextAlignment.Right;
			numberCell2.TextAnchorType = TextAnchorType.Bottom;
			var numberPortion2 = numberCell2.TextFrame.Paragraphs[0].Portions[0];
			numberPortion2.PortionFormat.FontHeight = 12;
			numberPortion2.PortionFormat.LatinFont = new FontData("Calibri");
			
			var numberCell3 = dataRow[4];
			numberCell3.TextFrame.Paragraphs[0].ParagraphFormat.Alignment = TextAlignment.Right;
			numberCell3.TextAnchorType = TextAnchorType.Bottom;
			var numberPortion3 = numberCell3.TextFrame.Paragraphs[0].Portions[0];
			numberPortion3.PortionFormat.FontHeight = 12;
		}
	}

	private async Task<string> CreateFilename()
	{
		var tempPath = Path.GetTempPath();
		var folder = Path.Combine(tempPath, "pptx-test");
		Directory.CreateDirectory(folder);
		var filename = Path.Combine(folder, $"{DateTimeOffset.UtcNow:yyyyMMddhhmmss}.pptx");
		return filename;
	}

	private void AddColumnChartSlide(int yearsCount, int monthsCount)
	{
		var slide = _presentationScope.Presentation.Slides.AddEmptySlide(_presentationScope.Presentation.LayoutSlides[0]);
		var slideSize = _presentationScope.Presentation.SlideSize.Size;
		
		SetSlideTitle(slide, "Столбчатая диаграмма");
		
		var chartX = 50f;
		var chartY = 150f;
		var chartWidth = (float)slideSize.Width - chartX * 2;
		var chartHeight = (float)slideSize.Height - chartY - 50f;
		
		var chart = slide.Shapes.AddChart(ChartType.ClusteredColumn, chartX, chartY, chartWidth, chartHeight);
		var workbook = chart.ChartData.ChartDataWorkbook;
		
		chart.ChartData.Series.Clear();
		chart.ChartData.Categories.Clear();
		
		var monthNames = new[] { "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь" };
		var categories = new string[monthsCount];
		for (int i = 0; i < monthsCount; i++)
		{
			categories[i] = monthNames[i % 12];
		}
		
		for (int i = 0; i < categories.Length; i++)
		{
			chart.ChartData.Categories.Add(workbook.GetCell(0, i + 1, 0, categories[i]));
		}
		
		var baseYear = 2023;
		var seriesList = new List<IChartSeries>();
		var colors = new[]
		{
			System.Drawing.Color.FromArgb(68, 114, 196),
			System.Drawing.Color.FromArgb(237, 125, 49),
			System.Drawing.Color.FromArgb(112, 173, 71),
			System.Drawing.Color.FromArgb(255, 192, 0),
			System.Drawing.Color.FromArgb(192, 0, 0)
		};
		
		for (int yearIndex = 0; yearIndex < yearsCount; yearIndex++)
		{
			var year = baseYear + yearIndex;
			var series = chart.ChartData.Series.Add(workbook.GetCell(0, 0, yearIndex + 1, $"Продажи {year}"), chart.Type);
			seriesList.Add(series);
			
			var random = new Random(year);
			for (int monthIndex = 0; monthIndex < monthsCount; monthIndex++)
			{
				var value = random.Next(30, 70);
				series.DataPoints.AddDataPointForBarSeries(workbook.GetCell(0, monthIndex + 1, yearIndex + 1, (double)value));
			}
			
			series.Format.Fill.FillType = FillType.Solid;
			series.Format.Fill.SolidFillColor.Color = colors[yearIndex % colors.Length];
		}
		
		chart.HasTitle = true;
		chart.ChartTitle.AddTextFrameForOverriding("Продажи по месяцам");
	}

	private void AddPieChartSlide(IReadOnlyDictionary<string, int> data)
	{
		var colors = new[]
		{
			System.Drawing.Color.FromArgb(68, 114, 196),
			System.Drawing.Color.FromArgb(237, 125, 49),
			System.Drawing.Color.FromArgb(112, 173, 71),
			System.Drawing.Color.FromArgb(255, 192, 0),
			System.Drawing.Color.FromArgb(192, 0, 0),
			System.Drawing.Color.FromArgb(112, 48, 160),
			System.Drawing.Color.FromArgb(0, 176, 240)
		};

		var slide = _presentationScope.Presentation.Slides.AddEmptySlide(_presentationScope.Presentation.LayoutSlides[0]);
		var slideSize = _presentationScope.Presentation.SlideSize.Size;
		
		SetSlideTitle(slide, "Круговая диаграмма");
		
		var chartX = 50f;
		var chartY = 150f;
		var chartWidth = (float)slideSize.Width - chartX * 2;
		var chartHeight = (float)slideSize.Height - chartY - 50f;
		
		var chart = slide.Shapes.AddChart(ChartType.Pie, chartX, chartY, chartWidth, chartHeight);
		var workbook = chart.ChartData.ChartDataWorkbook;
		
		chart.ChartData.Series.Clear();
		chart.ChartData.Categories.Clear();
		
		var items = data.ToList();
		var totalValue = items.Sum(x => x.Value);
		
		for (int i = 0; i < items.Count; i++)
		{
			chart.ChartData.Categories.Add(workbook.GetCell(0, i + 1, 0, items[i].Key));
		}
		
		var series = chart.ChartData.Series.Add(workbook.GetCell(0, 0, 1, "Доходы"), chart.Type);
		
		for (int i = 0; i < items.Count; i++)
		{
			var dataPoint = series.DataPoints.AddDataPointForPieSeries(workbook.GetCell(0, i + 1, 1, (double)items[i].Value));
			dataPoint.Format.Fill.FillType = FillType.Solid;
			dataPoint.Format.Fill.SolidFillColor.Color = colors[i % colors.Length];
		}
		
		var hasSmallSegments = items.Any(x => (double)x.Value / totalValue < 0.05);
		var canPlaceLabelsOnChart = items.Count <= 6 && !hasSmallSegments;
		
		if (canPlaceLabelsOnChart)
		{
			series.Labels.DefaultDataLabelFormat.ShowCategoryName = true;
			series.Labels.DefaultDataLabelFormat.ShowValue = false;
			series.Labels.DefaultDataLabelFormat.ShowLegendKey = false;
			series.Labels.DefaultDataLabelFormat.Position = LegendDataLabelPosition.BestFit;
			chart.Legend.Position = LegendPositionType.Bottom;
		}
		else
		{
			series.Labels.DefaultDataLabelFormat.ShowCategoryName = false;
			series.Labels.DefaultDataLabelFormat.ShowValue = false;
			series.Labels.DefaultDataLabelFormat.ShowLegendKey = false;
			chart.Legend.Position = LegendPositionType.Bottom;
		}
		
		chart.HasTitle = true;
		chart.ChartTitle.AddTextFrameForOverriding("Структура доходов");
	}

	private void AddLineChartSlide(int companiesCount, int quartersCount)
	{
		var slide = _presentationScope.Presentation.Slides.AddEmptySlide(_presentationScope.Presentation.LayoutSlides[0]);
		var slideSize = _presentationScope.Presentation.SlideSize.Size;
		
		SetSlideTitle(slide, "Линейный график");
		
		var chartX = 50f;
		var chartY = 150f;
		var chartWidth = (float)slideSize.Width - chartX * 2;
		var chartHeight = (float)slideSize.Height - chartY - 50f;
		
		var chart = slide.Shapes.AddChart(ChartType.LineWithMarkers, chartX, chartY, chartWidth, chartHeight);
		var workbook = chart.ChartData.ChartDataWorkbook;
		
		chart.ChartData.Series.Clear();
		chart.ChartData.Categories.Clear();
		
		var categories = new string[quartersCount];
		for (int i = 0; i < quartersCount; i++)
		{
			categories[i] = $"Q{i + 1}";
		}
		
		for (int i = 0; i < categories.Length; i++)
		{
			chart.ChartData.Categories.Add(workbook.GetCell(0, i + 1, 0, categories[i]));
		}
		
		var colors = new[]
		{
			System.Drawing.Color.FromArgb(68, 114, 196),
			System.Drawing.Color.FromArgb(237, 125, 49),
			System.Drawing.Color.FromArgb(112, 173, 71),
			System.Drawing.Color.FromArgb(255, 192, 0),
			System.Drawing.Color.FromArgb(192, 0, 0),
			System.Drawing.Color.FromArgb(112, 48, 160),
			System.Drawing.Color.FromArgb(0, 176, 240)
		};
		
		var random = new Random();
		
		for (int companyIndex = 0; companyIndex < companiesCount; companyIndex++)
		{
			var companyName = $"Компания {(char)('A' + companyIndex)}";
			var series = chart.ChartData.Series.Add(workbook.GetCell(0, 0, companyIndex + 1, companyName), chart.Type);
			
			var baseValue = random.Next(80, 150);
			for (int quarterIndex = 0; quarterIndex < quartersCount; quarterIndex++)
			{
				var value = baseValue + random.Next(-20, 30) * (quarterIndex + 1);
				if (value < 0) value = 0;
				series.DataPoints.AddDataPointForLineSeries(workbook.GetCell(0, quarterIndex + 1, companyIndex + 1, (double)value));
			}
			
			series.Format.Line.FillFormat.FillType = FillType.Solid;
			series.Format.Line.FillFormat.SolidFillColor.Color = colors[companyIndex % colors.Length];
		}
		
		chart.HasTitle = true;
		chart.ChartTitle.AddTextFrameForOverriding("Динамика роста");
	}

	private void AddAreaChartSlide(int regionsCount, int yearsCount)
	{
		var slide = _presentationScope.Presentation.Slides.AddEmptySlide(_presentationScope.Presentation.LayoutSlides[0]);
		var slideSize = _presentationScope.Presentation.SlideSize.Size;
		
		SetSlideTitle(slide, "Областная диаграмма");
		
		var chartX = 50f;
		var chartY = 150f;
		var chartWidth = (float)slideSize.Width - chartX * 2;
		var chartHeight = (float)slideSize.Height - chartY - 50f;
		
		var chart = slide.Shapes.AddChart(ChartType.Area, chartX, chartY, chartWidth, chartHeight);
		var workbook = chart.ChartData.ChartDataWorkbook;
		
		chart.ChartData.Series.Clear();
		chart.ChartData.Categories.Clear();
		
		var baseYear = 2020;
		var categories = new string[yearsCount];
		for (int i = 0; i < yearsCount; i++)
		{
			categories[i] = (baseYear + i).ToString();
		}
		
		for (int i = 0; i < categories.Length; i++)
		{
			chart.ChartData.Categories.Add(workbook.GetCell(0, i + 1, 0, categories[i]));
		}
		
		var colors = new[]
		{
			System.Drawing.Color.FromArgb(68, 114, 196),
			System.Drawing.Color.FromArgb(237, 125, 49),
			System.Drawing.Color.FromArgb(112, 173, 71),
			System.Drawing.Color.FromArgb(255, 192, 0),
			System.Drawing.Color.FromArgb(192, 0, 0),
			System.Drawing.Color.FromArgb(112, 48, 160),
			System.Drawing.Color.FromArgb(0, 176, 240)
		};
		
		var random = new Random();
		
		for (int regionIndex = 0; regionIndex < regionsCount; regionIndex++)
		{
			var regionName = $"Регион {regionIndex + 1}";
			var series = chart.ChartData.Series.Add(workbook.GetCell(0, 0, regionIndex + 1, regionName), chart.Type);
			
			var baseValue = random.Next(100, 200);
			for (int yearIndex = 0; yearIndex < yearsCount; yearIndex++)
			{
				var value = baseValue + random.Next(-20, 50) * (yearIndex + 1);
				if (value < 0) value = 0;
				series.DataPoints.AddDataPointForAreaSeries(workbook.GetCell(0, yearIndex + 1, regionIndex + 1, (double)value));
			}
			
			series.Format.Fill.FillType = FillType.Solid;
			series.Format.Fill.SolidFillColor.Color = colors[regionIndex % colors.Length];
		}
		
		chart.HasTitle = true;
		chart.ChartTitle.AddTextFrameForOverriding("Продажи по регионам");
	}

	private void AddScatterChartSlide()
	{
		var slide = _presentationScope.Presentation.Slides.AddEmptySlide(_presentationScope.Presentation.LayoutSlides[0]);
		var slideSize = _presentationScope.Presentation.SlideSize.Size;
		
		SetSlideTitle(slide, "Точечная диаграмма");
		
		var chartX = 50f;
		var chartY = 150f;
		var chartWidth = (float)slideSize.Width - chartX * 2;
		var chartHeight = (float)slideSize.Height - chartY - 50f;
		
		var chart = slide.Shapes.AddChart(ChartType.ScatterWithSmoothLinesAndMarkers, chartX, chartY, chartWidth, chartHeight);
		var workbook = chart.ChartData.ChartDataWorkbook;
		
		chart.ChartData.Series.Clear();
		
		var series = chart.ChartData.Series.Add(workbook.GetCell(0, 0, 1, "Зависимость"), chart.Type);
		
		var xValues = new[] { 1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0 };
		var yValues = new[] { 2.5, 5.1, 7.8, 10.2, 12.9, 15.3, 18.1, 20.5, 23.2, 25.8 };
		
		for (int i = 0; i < xValues.Length; i++)
		{
			series.DataPoints.AddDataPointForScatterSeries(
				workbook.GetCell(0, i + 1, 1, xValues[i]),
				workbook.GetCell(0, i + 1, 2, yValues[i]));
		}
		
		series.Format.Line.FillFormat.FillType = FillType.Solid;
		series.Format.Line.FillFormat.SolidFillColor.Color = System.Drawing.Color.FromArgb(68, 114, 196);
		
		chart.HasTitle = true;
		chart.ChartTitle.AddTextFrameForOverriding("Корреляция показателей");
	}

	private void AddCombinedChartSlide()
	{
		var slide = _presentationScope.Presentation.Slides.AddEmptySlide(_presentationScope.Presentation.LayoutSlides[0]);
		var slideSize = _presentationScope.Presentation.SlideSize.Size;
		
		SetSlideTitle(slide, "Сравнение диаграмм");
		
		var categories = new[] { "Продукт A", "Продукт B", "Продукт C", "Продукт D", "Продукт E" };
		var values = new[] { 42.0, 28.0, 18.0, 8.0, 4.0 };
		
		var chartMargin = 50f;
		var chartSpacing = 30f;
		var chartWidth = ((float)slideSize.Width - chartMargin * 2 - chartSpacing) / 2;
		var chartHeight = (float)slideSize.Height - 200f;
		var chartY = 150f;
		
		var pieChartX = chartMargin;
		var columnChartX = chartMargin + chartWidth + chartSpacing;
		
		var pieChart = slide.Shapes.AddChart(ChartType.Pie, pieChartX, chartY, chartWidth, chartHeight);
		var pieWorkbook = pieChart.ChartData.ChartDataWorkbook;
		
		pieChart.ChartData.Series.Clear();
		pieChart.ChartData.Categories.Clear();
		
		for (int i = 0; i < categories.Length; i++)
		{
			pieChart.ChartData.Categories.Add(pieWorkbook.GetCell(0, i + 1, 0, categories[i]));
		}
		
		var pieSeries = pieChart.ChartData.Series.Add(pieWorkbook.GetCell(0, 0, 1, "Продажи"), pieChart.Type);
		
		var pieColors = new[]
		{
			System.Drawing.Color.FromArgb(68, 114, 196),
			System.Drawing.Color.FromArgb(237, 125, 49),
			System.Drawing.Color.FromArgb(112, 173, 71),
			System.Drawing.Color.FromArgb(255, 192, 0),
			System.Drawing.Color.FromArgb(192, 0, 0)
		};
		
		for (int i = 0; i < values.Length; i++)
		{
			var dataPoint = pieSeries.DataPoints.AddDataPointForPieSeries(pieWorkbook.GetCell(0, i + 1, 1, values[i]));
			dataPoint.Format.Fill.FillType = FillType.Solid;
			dataPoint.Format.Fill.SolidFillColor.Color = pieColors[i];
		}
		
		pieChart.HasTitle = true;
		pieChart.ChartTitle.AddTextFrameForOverriding("Круговая диаграмма");
		
		var columnChart = slide.Shapes.AddChart(ChartType.ClusteredColumn, columnChartX, chartY, chartWidth, chartHeight);
		var columnWorkbook = columnChart.ChartData.ChartDataWorkbook;
		
		columnChart.ChartData.Series.Clear();
		columnChart.ChartData.Categories.Clear();
		
		for (int i = 0; i < categories.Length; i++)
		{
			columnChart.ChartData.Categories.Add(columnWorkbook.GetCell(0, i + 1, 0, categories[i]));
		}
		
		var columnSeries = columnChart.ChartData.Series.Add(columnWorkbook.GetCell(0, 0, 1, "Продажи"), columnChart.Type);
		
		for (int i = 0; i < values.Length; i++)
		{
			columnSeries.DataPoints.AddDataPointForBarSeries(columnWorkbook.GetCell(0, i + 1, 1, values[i]));
		}
		
		columnSeries.Format.Fill.FillType = FillType.Solid;
		columnSeries.Format.Fill.SolidFillColor.Color = System.Drawing.Color.FromArgb(68, 114, 196);
		
		columnChart.HasTitle = true;
		columnChart.ChartTitle.AddTextFrameForOverriding("Столбчатая диаграмма");
	}

	private void SetSlideTitle(ISlide slide, string title)
	{
		IAutoShape? titleShape = null;
		foreach (IShape shape in slide.Shapes)
		{
			if (shape is IAutoShape autoShape && autoShape.TextFrame != null)
			{
				if (shape.Placeholder != null && 
				    (shape.Placeholder.Type == PlaceholderType.Title || shape.Placeholder.Type == PlaceholderType.CenteredTitle))
				{
					titleShape = autoShape;
					break;
				}
			}
		}
		
		if (titleShape == null)
		{
			foreach (IShape shape in slide.Shapes)
			{
				if (shape is IAutoShape autoShape && autoShape.TextFrame != null && shape.Y < 200)
				{
					titleShape = autoShape;
					break;
				}
			}
		}
		
		if (titleShape != null)
		{
			titleShape.TextFrame.Text = title;
		}
	}

	private async Task SavePresentation(string filename)
	{
		_presentationScope.Presentation.Save(filename, SaveFormat.Pptx);
	}
}