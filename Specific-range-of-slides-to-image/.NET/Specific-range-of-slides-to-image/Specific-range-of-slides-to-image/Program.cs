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
                    // Loop through a specific range of slides and convert them to images.
                    // The loop processes slides starting from the 3rd slide (index 2) to the second-to-last slide.
                    for (int currentSlideIndex = 2; currentSlideIndex < imageStreams.Length-1; currentSlideIndex++)
                    {
                        //Convert the PowerPoint slide as an image stream.
                        using (Stream stream = pptxDoc.Slides[currentSlideIndex].ConvertToImage(ExportImageFormat.Jpeg))
                        {
                            //Save the image stream to a file.
                            using (FileStream fileStreamOutput = File.Create(Path.GetFullPath(@"../../../Image_" + currentSlideIndex + ".jpeg")))
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
