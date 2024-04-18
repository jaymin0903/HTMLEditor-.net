using DocumentFormat.OpenXml.InkML;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.AspNetCore.Components.Forms;

//using Google.Apis.Auth.OAuth2;
//using Google.Apis.Services;
//using Microsoft.AspNetCore.Cors.Infrastructure;
using Microsoft.AspNetCore.Mvc;
using OpenXmlPowerTools;
using Org.BouncyCastle.Asn1.Pkcs;

//using Org.BouncyCastle.Asn1.Ocsp;
using System.Xml.Linq;
//using Google.Apis.Auth.OAuth2;
//using Google.Apis.Services;
//using Google.Apis.Util.Store;
//using Google.Apis.Docs.v1;
//using System.Collections.Generic;
//using System.IO;
//using System.Threading;
//using System.Threading.Tasks;
//using Google.Apis.Docs.v1.Data;

namespace HTMLEditor.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class DocxToHtmlController : ControllerBase
    {
        private readonly ILogger<DocxToHtmlController> _logger;
        private readonly EditorContext _context;
        public DocxToHtmlController(ILogger<DocxToHtmlController> logger, EditorContext context)
        {
            _logger = logger;
            _context = context;
        }
        [HttpPost("convert")]
        public async Task<ContentResult> ConvertDocxToHtml()
        {
            var file = Request.Form.Files[0];

            if (file.Length > 0)
            {
                using (var stream = new MemoryStream())
                {
                    await file.CopyToAsync(stream);
                    using (WordprocessingDocument doc = WordprocessingDocument.Open(stream, true))
                    {
                        WmlToHtmlConverterSettings settings = new()
                        {
                            PageTitle = "Title"
                        };
                        XElement html = WmlToHtmlConverter.ConvertToHtml(doc, settings);

                        string bodyHtml = html.ToString();

                        return new ContentResult
                        {
                            ContentType = "text/html",
                            Content = bodyHtml,
                            StatusCode = 200
                        };
                    }
                }
            }
            return new ContentResult
            {
                StatusCode = 500
            };
        }

        [HttpPost("save")]
        public async Task<ActionResult> StoreDoc([FromForm] string htmlData) {
            try
            {
                var docData = new FileContent
                {
                    FileName = "Document Name", // You may want to pass this in the request
                    FileText = htmlData
                };

                _context.FileContent.Add(docData);
                await _context.SaveChangesAsync();

                return Ok("Data saved successfully");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error saving data");
                return StatusCode(500, "Internal server error");
            }
        }

        #region ConvertDocxToHtmlUsingGoogleAPI
        //public async Task<string> ConvertDocxToHtmlUsingGoogleAPI(Stream docxFileStream)
        //{
        //    // Authenticate using service account credentials
        //    GoogleCredential credential = GoogleCredential.FromFile("C:\\Users\\sit374.SIT\\Documents\\Jaymin\\.NET\\Web API\\HTMLEditor\\HTMLEditor\\my-project-1703157798408-3dd702c952fc.json")
        //        .CreateScoped(DocsService.ScopeConstants.Documents);

        //    // Create the DocsService using the service account credentials
        //    var service = new DocsService(new BaseClientService.Initializer()
        //    {
        //        HttpClientInitializer = credential,
        //        ApplicationName = "html editor",
        //    });

        //    // Create the request to upload the document
        //    // Upload the document
        //    var newDocument = new Google.Apis.Docs.v1.Data.Document();
        //    var uploadRequest = service.Documents.Create(newDocument);
        //    var uploadedDocument = await uploadRequest.ExecuteAsync();

        //    // Update the document content
        //    var updateRequest = service.Documents.BatchUpdate(new BatchUpdateDocumentRequest
        //    {
        //        Requests = new List<Google.Apis.Docs.v1.Data.Request>
        //{
        //    new Google.Apis.Docs.v1.Data.Request
        //    {
        //        UpdateDocumentStyle = new UpdateDocumentStyleRequest
        //        {
        //            DocumentStyle = new DocumentStyle
        //            {
        //                BackgroundColor = new OptionalColor
        //                {
        //                    Color = new Color
        //                    {
        //                        RgbColor = new RgbColor
        //                        {
        //                            Red = 1.0f,
        //                            Green = 1.0f,
        //                            Blue = 1.0f
        //                        }
        //                    }
        //                }
        //            },
        //            Fields = "backgroundColor"
        //        }
        //    },
        //    new Google.Apis.Docs.v1.Data.Request
        //    {
        //        InsertText = new InsertTextRequest
        //        {
        //            Text = "Your document content here.",
        //            Location = new Location
        //            {
        //                Index = 1
        //            }
        //        }
        //    }
        //}
        //    }, uploadedDocument.DocumentId);

        //    await updateRequest.ExecuteAsync();

        //    // Export the document as HTML
        //    var exportRequest = service.Documents.Export(uploadedDocument.DocumentId, "text/html");
        //    var htmlContent = await exportRequest.ExecuteAsync();

        //    return htmlContent;
        //}
        #endregion

        #region ConvertWorkbookToHtml
        //private string ConvertWorkbookToHtml(HSSFWorkbook workbook)
        //{
        //    var html = new StringBuilder();
        //    var document = new StreamDocument();

        //    foreach (ISheet sheet in workbook.Sheets)
        //    {
        //        html.Append("<h2>Sheet " + sheet.SheetName + "</h2>");
        //        html.Append("<table>");

        //        var rowIndex = 0;
        //        foreach (IRow row in sheet)
        //        {
        //            if (rowIndex == 0)
        //            {
        //                html.Append("<thead>");
        //                foreach (ICell cell in row)
        //                {
        //                    html.Append("<th>" + cell.StringCellValue + "</th>");
        //                }
        //                html.Append("</thead>");
        //            }
        //            else
        //            {
        //                html.Append("<tbody>");
        //                foreach (ICell cell in row)
        //                {
        //                    html.Append("<td>" + cell.StringCellValue + "</td>");
        //                }
        //                html.Append("</tbody>");
        //            }

        //            rowIndex++;
        //        }

        //        html.Append("</table>");
        //    }

        //    return html.ToString();
        //}
        #endregion

    }
}
