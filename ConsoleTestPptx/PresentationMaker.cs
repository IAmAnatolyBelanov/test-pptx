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