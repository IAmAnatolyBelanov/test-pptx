using Aspose.Slides;
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
		AddTableSlide();
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
		thirdPortion.PortionFormat.FillFormat.FillType = FillType.Solid;
		thirdPortion.PortionFormat.FillFormat.SolidFillColor.Color = System.Drawing.Color.Black;
	}

	private void AddTableSlide()
	{
		var slide = _presentationScope.Presentation.Slides.AddEmptySlide(_presentationScope.Presentation.LayoutSlides[0]);
		var slideSize = _presentationScope.Presentation.SlideSize.Size;
		
		var titleWidth = 600f;
		var titleHeight = 60f;
		var titleX = ((float)slideSize.Width - titleWidth) / 2;
		var titleY = 50f;
		
		var titleBox = slide.Shapes.AddAutoShape(ShapeType.Rectangle, titleX, titleY, titleWidth, titleHeight);
		titleBox.FillFormat.FillType = FillType.NoFill;
		titleBox.LineFormat.FillFormat.FillType = FillType.NoFill;
		titleBox.TextFrame.Text = "название таблицы";
		titleBox.TextFrame.Paragraphs[0].ParagraphFormat.Alignment = TextAlignment.Center;
		var titlePortion = titleBox.TextFrame.Paragraphs[0].Portions[0];
		titlePortion.PortionFormat.FontHeight = 32;
		titlePortion.PortionFormat.FillFormat.FillType = FillType.Solid;
		titlePortion.PortionFormat.FillFormat.SolidFillColor.Color = System.Drawing.Color.Black;
		
		var tableX = ((float)slideSize.Width - 900) / 2;
		var tableY = titleY + titleHeight + 40f;
		
		double[] columnWidths = { 180, 180, 180, 180, 180 };
		double[] rowHeights = { 50, 50, 50, 50, 50 };
		
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
		
		for (int row = 1; row < 5; row++)
		{
			var dataRow = table.Rows[row];
			dataRow[0].TextFrame.Text = $"Текст {row}";
			dataRow[1].TextFrame.Text = DateTime.Now.AddDays(row).ToString("dd.MM.yyyy HH:mm");
			dataRow[2].TextFrame.Text = (row * 10).ToString();
			dataRow[3].TextFrame.Text = (row * 20).ToString();
			dataRow[4].TextFrame.Text = (row * 30).ToString();
			
			for (int col = 0; col < 5; col++)
			{
				var cell = dataRow[col];
				cell.TextFrame.Paragraphs[0].ParagraphFormat.Alignment = col == 0 ? TextAlignment.Left : TextAlignment.Center;
				var portion = cell.TextFrame.Paragraphs[0].Portions[0];
				portion.PortionFormat.FontHeight = 12;
			}
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

	private async Task SavePresentation(string filename)
	{
		_presentationScope.Presentation.Save(filename, SaveFormat.Pptx);
	}
}