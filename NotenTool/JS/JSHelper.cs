using System.Text.Json;
using Blazor.DownloadFileFast.Interfaces;
using Microsoft.JSInterop;
using NotenTool.DTO;

namespace NotenTool.JS;

public class JsHelper(IJSRuntime js, IBlazorDownloadFileService downloadFileService)
{
    public async Task DownloadFileAsync(byte[] data, string fileName)
    {
        await downloadFileService.DownloadFileAsync(fileName, data);
    }
    
    public async Task<FilePickerResult?> PickFileAsync(string accept = ".xlsx")
    {
        try
        {
            var result = await js.InvokeAsync<FilePickerResult>("pickFile", accept);
            result.UploadDate = DateTime.UtcNow;
            return result;
        }
        catch
        {
            return null;
        }
    }
    
    public async Task SetItemToLocalStorageAsync<T>(T value, string key ="cachedCourseData")
    {
        var json = JsonSerializer.Serialize(value);
        await js.InvokeVoidAsync("localStorage.setItem", key, json);
    }

    public async Task<T?> GetItemFromLocalStorageAsync<T>(string key = "cachedCourseData")
    {
        var json = await js.InvokeAsync<string>("localStorage.getItem", key);
        
        if (string.IsNullOrEmpty(json))
            return default;
            
        return JsonSerializer.Deserialize<T>(json);
    }

    public async Task RemoveItemFromLocalStorageAsync(string key)
    {
        await js.InvokeVoidAsync("localStorage.removeItem", key);
    }

    public async Task ClearLocalStorageAsync()
    {
        await js.InvokeVoidAsync("localStorage.clear");
    }
}