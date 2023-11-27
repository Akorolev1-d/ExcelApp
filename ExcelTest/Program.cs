//using ExcelTest.Data;
using Microsoft.EntityFrameworkCore;
using NPOI.HPSF;
using NPOI.SS.Formula.Functions;
using System;
using System.IO;
internal class Program
{
    private static void Main(string[] args)
    {
        var builder = WebApplication.CreateBuilder(args);

        var connectionString = builder.Configuration.GetConnectionString("DefaultConnection") ?? throw new InvalidOperationException("Connection String 'DefaultConnection' not found");
        //builder.Services.AddDbContext<AppicationDbContext>(options => options.UseSqlServer(connectionString));

        // Add services to the container.
        builder.Services.AddControllersWithViews();
        builder.Services.AddRazorPages().AddRazorRuntimeCompilation();

        string path = "Uploads";
        string fullPath = Path.GetFullPath(path); // "C:\\Users\\PK\\source\\repos\\ExcelTest\\ExcelTest\\Uploads"
        int cc = fullPath.Count() - 1;
        fullPath = fullPath.Remove(Path.GetFullPath(path).Count() - 8, 8) + "\\wwwroot" + "\\Uploads";

        DirectoryInfo di = new DirectoryInfo(fullPath);
        foreach (var fi in di.GetFiles())
        {
            try { fi.Delete(); } catch { }
        }
        try { di.Delete(); } catch { }
        di.Create();


        var app = builder.Build();

        // Configure the HTTP request pipeline.
        if (!app.Environment.IsDevelopment())
        {
            app.UseExceptionHandler("/Home/Error");
            // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
            app.UseHsts();
        }

        app.UseHttpsRedirection();
        app.UseStaticFiles();

        app.UseRouting();

        app.UseAuthorization();

        app.MapControllerRoute(
            name: "default",
            pattern: "{controller=Home}/{action=Index}/{id?}");

        app.Run();
    }
}