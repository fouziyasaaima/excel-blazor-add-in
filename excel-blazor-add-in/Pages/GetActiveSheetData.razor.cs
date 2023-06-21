using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace BlazorAddIn.Pages
{
    public partial class GetActiveSheetData
    {
        [Inject]
        public IJSRuntime JSRuntime { get; set; } = default!;
        public IJSObjectReference JSModule { get; set; } = default!;

        public string? ExcelData { get; set; }


        protected override async Task OnAfterRenderAsync(bool firstRender)
        {
            if (firstRender)
            {
                JSModule = await JSRuntime.InvokeAsync<IJSObjectReference>("import", "./Pages/GetActiveSheetData.razor.js");
            }
        }

        /// <summary>
        /// Function to get Data from Excel
        /// </summary>
        private async Task RetrieveActiveSheetData()
        {

            ExcelData = await JSModule.InvokeAsync<String>("GetData");
            StateHasChanged();
       
        }
              
        //Event CallBack from Javascript to Excel 
        [JSInvokable]
        public void GetExcelData(string data)
        {
            ExcelData = data;
            StateHasChanged();
        }

    }

}

