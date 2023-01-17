using Aspose.Words;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using FileManage.Interface;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.StaticFiles;
using System.Net.Mime;
using System.Text;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace FileManage.Controllers
{
    [ApiController]
    [Route("api")]
    public class FileController : Controller
    {
        private readonly IPhotoService _photoService;
        public FileController(IPhotoService photoService)
        {
            _photoService = photoService;
        }

        [HttpPost]
        [Route("UploadFile")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(typeof(string), StatusCodes.Status400BadRequest)]
        public async Task<IActionResult> UploadFile(IFormFile file, CancellationToken cancellationtoken)
        {
            var result = await WriteFile(file, null);
            return Ok(result);
        }

        private async Task<string> WriteFile(IFormFile file, string type)
        {
            string filename = "";
            string filePathName = "Upload\\Files";
            if (type == "PDF")
            {
                filePathName = "Upload\\Temp\\PDF";
            }
            else if(type == "IMAGE")
            {
                filePathName = "Upload\\Temp\\IMAGE";
            }
            try
            {
                var extension = "." + file.FileName.Split('.')[file.FileName.Split('.').Length - 1];
                filename = DateTime.Now.Ticks.ToString() + extension;

                var filepath = Path.Combine(Directory.GetCurrentDirectory(), filePathName);

                if (!Directory.Exists(filepath))
                {
                    Directory.CreateDirectory(filepath);
                }

                var exactpath = Path.Combine(Directory.GetCurrentDirectory(), filePathName, filename);
                using (var stream = new FileStream(exactpath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }
            }
            catch (Exception ex)
            {
            }
            return filename;
        }

        [HttpGet]
        [Route("DownloadFile")]
        public async Task<IActionResult> DownloadFile(string filename)
        {
            string filepath = Path.Combine(Directory.GetCurrentDirectory(), "Upload\\Files", filename);

            FileExtensionContentTypeProvider provider = new FileExtensionContentTypeProvider();
            if (!provider.TryGetContentType(filepath, out var contenttype))
            {
                contenttype = "application/octet-stream";
            }

            byte[] bytes = await System.IO.File.ReadAllBytesAsync(filepath);
            FileContentResult result = File(bytes, contenttype, Path.GetFileName(filepath));
            return result;
        }

        [HttpPost("byte")]
        public async Task<IActionResult> ConvertToByte(IFormFile formFile)
        {
            var stream = new MemoryStream((int)formFile.Length);
            formFile.CopyTo(stream);
            var bytes = stream.ToArray();
            FileContentResult result = File(bytes, "application/octet-stream", formFile.FileName);
            return result;
        }
        [HttpPost("base64")]
        public async Task<IActionResult> ConvertToBase64(IFormFile formFile)
        {
            string base64 = string.Empty;
            if (formFile.Length > 0)
            {
                using (var ms = new MemoryStream())
                {
                    formFile.CopyTo(ms);
                    var fileBytes = ms.ToArray();
                    base64 = Convert.ToBase64String(fileBytes);                  
                }
            }
            byte[] bytes = Convert.FromBase64String(base64);
            string base64String = Convert.ToBase64String(bytes, 0, bytes.Length); // for bytes to base64string
            FileContentResult result = File(bytes, "application/octet-stream", formFile.FileName);
            return result;
        }
        private class User
        {
            public int Id { get; set; }
            public string Username { get; set; }
        }

        private List<User> users = new List<User>
    {
        new User { Id = 1, Username = "DoloresAbernathy" },
        new User { Id = 2, Username = "MaeveMillay" },
        new User { Id = 3, Username = "BernardLowe" },
        new User { Id = 4, Username = "ManInBlack" }
    };      

        [HttpGet("ExelSheet")]
        public async Task<IActionResult> Sheet()
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Users");
                var currentRow = 1;
                worksheet.Cell(currentRow, 1).Value = "Id";
                worksheet.Cell(currentRow, 2).Value = "Username";
                foreach (var user in users)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = user.Id;
                    worksheet.Cell(currentRow, 2).Value = user.Username;
                }
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();

                    return File(
                        content,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "users.xlsx");
                }
            }
        }

        [HttpGet("CVS")]
        public async Task<IActionResult> CVS()
        {
            var builder = new StringBuilder();
            builder.AppendLine("Id,Username");
            foreach (var user in users)
            {
                builder.AppendLine($"{user.Id},{user.Username}");
            }

            return File(Encoding.UTF8.GetBytes(builder.ToString()), "text/csv", "users.csv");
        }

        [HttpPost("pdfToWord")]
        public async Task<IActionResult> PdfToWordConverter(IFormFile formFile)
        {
            string filename = await WriteFile(formFile, "PDF");
            string pdfFilepath = Path.Combine(Directory.GetCurrentDirectory(), "Upload\\Temp\\PDF", filename);
            var doc = new Document(pdfFilepath);
            string[] strings = formFile.FileName.Split('.');
            string docFilename = $"{formFile.FileName.Replace($".{strings.Last()}", "")}.docx";
            doc.Save(Path.Combine(Directory.GetCurrentDirectory(), $"Upload\\Temp\\WORD\\{docFilename}"));  
            
            string filepath = Path.Combine(Directory.GetCurrentDirectory(), "Upload\\Temp\\WORD", docFilename);
            FileExtensionContentTypeProvider provider = new FileExtensionContentTypeProvider();
            if (!provider.TryGetContentType(filepath, out var contenttype))
            {
                contenttype = "application/octet-stream";
            }
            byte[] bytes = await System.IO.File.ReadAllBytesAsync(filepath);
            FileContentResult result = File(bytes, contenttype, Path.GetFileName(filepath));
            System.IO.File.Delete(pdfFilepath);
            System.IO.File.Delete(filepath);          
           
            return result;
        }
        //https://products.aspose.com/words/net/conversion/pdf-to-image/ use this site for other formate conversion 
        [HttpPost("pdfToImage")]
        public async Task<IActionResult> PdfToImage(IFormFile formFile)
        {
            string filename = await WriteFile(formFile, "PDF");
            string pdfFilepath = Path.Combine(Directory.GetCurrentDirectory(), "Upload\\Temp\\PDF", filename);
            var doc = new Document(pdfFilepath);
            string[] strings = formFile.FileName.Split('.');
            for (int page = 0; page < doc.PageCount; page++)
            {
                var extractedPage = doc.ExtractPages(page, 1);
                extractedPage.Save($"Upload\\Temp\\IMAGE\\{formFile.FileName.Replace($".{strings.Last()}", "")}_{page + 1}.jpg");
            }
            System.IO.File.Delete(pdfFilepath);
            return Ok();
        }
        [HttpPost("uploadImage")]
        public async Task<IActionResult> UploadImageToCloudinary(IFormFile formFile)
        {
            return Ok(await _photoService.AddPhotoAsync(formFile));
        }
        [HttpGet("deleteImage")]
        public async Task<IActionResult> DeleteImageToCloudinary([FromQuery] string publicId)
        {
            return Ok(await _photoService.DeletePhotoAsync(publicId));
        }
    }
}
