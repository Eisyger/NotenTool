using Application.Services;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.AspNetCore.Components.WebAssembly.Hosting;
using NotenTool;
using NotenTool.JS;

var builder = WebAssemblyHostBuilder.CreateDefault(args);
builder.RootComponents.Add<App>("#app");
builder.RootComponents.Add<HeadOutlet>("head::after");

builder.Services.AddScoped(sp => new HttpClient { BaseAddress = new Uri(builder.HostEnvironment.BaseAddress) });
builder.Services.AddScoped<ExcelTemplateExportService>();
builder.Services.AddScoped<JsHelper>();
builder.Services.AddScoped<ExcelParseService>();
builder.Services.AddScoped<RandomService>();
builder.Services.AddScoped<ExcelGradesExportService>();

await builder.Build().RunAsync();