using Syncfusion.Drawing;
using Syncfusion.OfficeChartToImageConverter;
using Syncfusion.Presentation;
using System.IO;

namespace convert_powerpoint_slide_to_image
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Open the  file as Stream.
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"../../Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Load an existing PowerPoint presentation. 
                using (IPresentation pptxDoc = Presentation.Open(inputStream))
                {
                    //Initialize ChartToImageConverter. 
                    pptxDoc.ChartToImageConverter = new ChartToImageConverter();
                    //Convert the first slide of the PowerPoint to an image.
                    using (Stream stream = pptxDoc.Slides[0].ConvertToImage((ImageFormat)ImageType.Metafile))
                    {
                        //Save the image stream to file.
                        using (FileStream fileStreamOutput = File.Create(Path.GetFullPath(@"../../Image.jpg")))
                        {
                            stream.CopyTo(fileStreamOutput);
                        }
                    }
                }
            }
        }
    }
}
