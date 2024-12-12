using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

namespace Convert_PowerPoint_slide_to_Image
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Open an existing presentation.
            using (FileStream inputStream = new(Path.GetFullPath(@"../../../Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Load the presentation. 
                using (IPresentation pptxDoc = Presentation.Open(inputStream))
                {
                    //Initialize PresentationRenderer. 
                    pptxDoc.PresentationRenderer = new PresentationRenderer();
                    //Convert the PowerPoint 2nd slide as an image stream.
                    using (Stream stream = pptxDoc.Slides[1].ConvertToImage(ExportImageFormat.Jpeg))
                    {
                        //Save the image stream to a file.
                        using (FileStream fileStreamOutput = File.Create(Path.GetFullPath(@"../../../Image.jpg")))
                        {
                            stream.CopyTo(fileStreamOutput);
                        }
                    }
                }
            }
        }
    }
}