using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

namespace convert_specific_range_of_slides_to_images
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Open the  file as Stream.
            using (FileStream inputStream = new(Path.GetFullPath(@"../../../Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Load an existing PowerPoint presentation.
                using (IPresentation pptxDoc = Presentation.Open(inputStream))
                {
                    //Initialize PresentationRenderer. 
                    pptxDoc.PresentationRenderer = new PresentationRenderer();
                    //Loop through a specific range of slides and convert each to an image.
                    for (int currentSlideIndex = 2; currentSlideIndex < pptxDoc.Slides.Count-1; currentSlideIndex++)
                    {
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
