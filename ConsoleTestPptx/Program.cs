using Aspose.Slides;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using Microsoft.Extensions.Hosting;

namespace ConsoleTestPptx;

public class Program
{
	public static async Task Main(string[] args)
	{
		LoadLicenseIfExists();

		var builder = Host.CreateApplicationBuilder(args);
		builder.Services.TryAddScoped<PresentationScope>();
		builder.Services.TryAddScoped<PresentationMaker>();

		var host = builder.Build();

		await using var scope = host.Services.CreateAsyncScope();
		var presentationScope = scope.ServiceProvider.GetRequiredService<PresentationScope>();
		presentationScope.Presentation = new Presentation();
		// presentationScope.Presentation.SlideSize.SetSize(1920f, 1080f, SlideSizeScaleType.DoNotScale);
		var maker = scope.ServiceProvider.GetRequiredService<PresentationMaker>();
		await maker.BuildPresentation();
	}

	private static void LoadLicenseIfExists()
	{
		var baseDirectory = AppContext.BaseDirectory;
		var licenseFiles = new[] { "Aspose.Slides.lic", "license.lic", "Aspose.Slides.lic.xml" };

		foreach (var licenseFile in licenseFiles)
		{
			var licensePath = Path.Combine(baseDirectory, licenseFile);
			if (File.Exists(licensePath))
			{
				var license = new License();
				license.SetLicense(licensePath);
				return;
			}
		}
	}
}