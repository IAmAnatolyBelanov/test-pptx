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
		AddTitleToFirstSlide();
		var filename = await CreateFilename();
		await SavePresentation(filename);
	}

	private void AddTitleToFirstSlide()
	{
		var slide = _presentationScope.Presentation.Slides.Count > 0
			? _presentationScope.Presentation.Slides[0]
			: _presentationScope.Presentation.Slides.AddEmptySlide(_presentationScope.Presentation.LayoutSlides[0]);

		var slideSize = _presentationScope.Presentation.SlideSize.Size;
		var textBox = slide.Shapes.AddTextBox(ShapeType.Rectangle, 0, 0, (float)slideSize.Width, (float)slideSize.Height);
		textBox.TextFrame.Text = "Тестовый тайтл";
		textBox.TextFrame.Paragraphs[0].ParagraphFormat.Alignment = TextAlignment.Center;
		textBox.TextFrame.Paragraphs[0].Portions[0].PortionFormat.FontHeight = 48;
		
		var textFrame = textBox.TextFrame;
		textFrame.AnchoringType = TextAnchorType.Center;
		textFrame.VerticalAnchorType = TextVerticalType.Center;
		textFrame.AutofitType = TextAutofitType.None;
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