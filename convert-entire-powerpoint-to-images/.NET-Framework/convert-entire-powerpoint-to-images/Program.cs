using Syncfusion.Drawing;
using Syncfusion.OfficeChartToImageConverter;
using Syncfusion.Presentation;
using System.IO;

namespace convert_entire_powerpoint_to_images
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Open the file as Stream.
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../Template.pptx"), FileMode.Open, FileAccess.Read))
            {
                //Load an existing PowerPoint presentation.
                using (IPresentation pptxDoc = Presentation.Open(fileStream))
                {
                    //Initialize ChartToImageConverter
                    pptxDoc.ChartToImageConverter = new ChartToImageConverter();
                    //Convert PowerPoint to images.
                    Stream[] images = pptxDoc.RenderAsImages((ImageFormat)ImageType.Metafile);
                    //Save the images to file.
                    for (int i = 0; i < images.Length; i++)
                    {
                        using (Stream stream = images[i])
                        {
                            using (FileStream fileStreamOutput = File.Create(Path.GetFullPath("../../Image-" + i + ".jpg")))
                            {
                                stream.CopyTo(fileStreamOutput);
                            }
                        }
                    }
                }
            }
        }
    }
}
