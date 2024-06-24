using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.IO;
using System.Linq;

class PowerPointSplitter
{
    static void Main(string[] args)
    {
        string inputFile = @"path\to\your\input.pptx";
        string outputFolder = @"path\to\output\folder";

        SplitPowerPoint(inputFile, outputFolder);
    }

    static void SplitPowerPoint(string inputFile, string outputFolder)
    {
        using (PresentationDocument presentationDocument = PresentationDocument.Open(inputFile, false))
        {
            var presentation = presentationDocument.PresentationPart.Presentation;
            var slideIdList = presentation.SlideIdList;

            for (int i = 0; i < slideIdList.ChildElements.Count; i++)
            {
                string newPresentationPath = Path.Combine(outputFolder, $"Slide_{i + 1}.pptx");
                File.Copy(inputFile, newPresentationPath, true);

                using (PresentationDocument newPresentation = PresentationDocument.Open(newPresentationPath, true))
                {
                    var newPresSlideIdList = newPresentation.PresentationPart.Presentation.SlideIdList;

                    // Remove all slides except the one we want to keep
                    for (int j = newPresSlideIdList.ChildElements.Count - 1; j >= 0; j--)
                    {
                        if (j != i)
                        {
                            var slideId = (SlideId)newPresSlideIdList.ChildElements[j];
                            var slidePart = newPresentation.PresentationPart.GetPartById(slideId.RelationshipId) as SlidePart;
                            newPresSlideIdList.RemoveChild(slideId);
                            newPresentation.PresentationPart.DeletePart(slidePart);
                        }
                    }

                    // Update slide numbers
                    newPresentation.PresentationPart.Presentation.Save();
                }
            }
        }

        Console.WriteLine("PowerPoint split completed successfully.");
    }
}
