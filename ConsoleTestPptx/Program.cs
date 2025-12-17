using Aspose.Slides;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using Microsoft.Extensions.Hosting;

namespace ConsoleTestPptx;

public class Program
{
	public static async Task Main(string[] args)
	{
		var builder = Host.CreateApplicationBuilder(args);
		builder.Services.TryAddScoped<PresentationScope>();
		builder.Services.TryAddScoped<PresentationMaker>();

		var host = builder.Build();

		await using var scope = host.Services.CreateAsyncScope();
		var presentationScope = scope.ServiceProvider.GetRequiredService<PresentationScope>();
		presentationScope.Presentation = new Presentation();
		var maker = scope.ServiceProvider.GetRequiredService<PresentationMaker>();
		await maker.BuildPresentation();
	}
}