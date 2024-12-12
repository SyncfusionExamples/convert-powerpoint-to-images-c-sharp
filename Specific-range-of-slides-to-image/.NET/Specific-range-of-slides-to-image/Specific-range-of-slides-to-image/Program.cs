using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

namespace Specific_range_of_slides_to_image
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
                    //Convert PowerPoint to image as stream.
                    Stream[] imageStreams = pptxDoc.RenderAsImages(ExportImageFormat.Jpeg);
                    // Loop through the slides, converting only a subset of slides to images.
                    // 'Length - 1' ensures only a specific range of slides (e.g., the first n-1 slides) are converted.
                    // Modify the range logic based on requirements to include or exclude specific slides.
                    for (int i = 0; i < imageStreams.Length-1; i++)
                    {
                        //Convert the PowerPoint slide as an image stream.
                        using (Stream stream = pptxDoc.Slides[0].ConvertToImage(ExportImageFormat.Jpeg))
                        {
                            //Save the image stream to a file.
                            using (FileStream fileStreamOutput = File.Create(Path.GetFullPath(@"../../../Image_" + i + ".jpeg")))
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
