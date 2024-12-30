using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

namespace Create_PowerPoint_presentation
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Open the file as Stream.
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Template.pptx"), FileMode.Open, FileAccess.Read))
            {
                //Load an existing PowerPoint presentation.
                using (IPresentation pptxDoc = Presentation.Open(fileStream))
                {
                    //Initialize PresentationRenderer.
                    pptxDoc.PresentationRenderer = new PresentationRenderer();
                    //Convert PowerPoint to images.
                    Stream[] images = pptxDoc.RenderAsImages(ExportImageFormat.Jpeg);
                    //Save the images to file.
                    for (int i = 0; i < images.Length; i++)
                    {
                        using (Stream stream = images[i])
                        {
                            using (FileStream fileStreamOutput = File.Create(Path.GetFullPath("../../../Image-" + i + ".jpg")))
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