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
                
                using (PresentationDocument newPresentation = PresentationDocument.Create(newPresentationPath, DocumentFormat.OpenXml.PresentationDocumentType.Presentation))
                {
                    var newPresentationPart = newPresentation.PresentationPart;
                    if (newPresentationPart == null)
                    {
                        newPresentationPart = newPresentation.AddPresentationPart();
                    }

                    newPresentationPart.Presentation = new Presentation();

                    // Copy slide
                    var slideId = (SlideId)slideIdList.ChildElements[i];
                    var sourceSlidePart = (SlidePart)presentationDocument.PresentationPart.GetPartById(slideId.RelationshipId);
                    var targetSlidePart = newPresentationPart.AddNewPart<SlidePart>();
                    
                    using (Stream sourceStream = sourceSlidePart.GetStream())
                    using (Stream targetStream = targetSlidePart.GetStream())
                    {
                        sourceStream.CopyTo(targetStream);
                    }

                    // Copy slide layout
                    var sourceSlideLayoutPart = sourceSlidePart.SlideLayoutPart;
                    var targetSlideLayoutPart = targetSlidePart.AddNewPart<SlideLayoutPart>();
                    
                    using (Stream sourceStream = sourceSlideLayoutPart.GetStream())
                    using (Stream targetStream = targetSlideLayoutPart.GetStream())
                    {
                        sourceStream.CopyTo(targetStream);
                    }

                    // Copy slide master
                    var sourceSlideMasterPart = sourceSlideLayoutPart.SlideMasterPart;
                    var targetSlideMasterPart = targetSlideLayoutPart.AddNewPart<SlideMasterPart>();
                    
                    using (Stream sourceStream = sourceSlideMasterPart.GetStream())
                    using (Stream targetStream = targetSlideMasterPart.GetStream())
                    {
                        sourceStream.CopyTo(targetStream);
                    }

                    // Copy theme
                    var sourceThemePart = sourceSlideMasterPart.ThemePart;
                    var targetThemePart = targetSlideMasterPart.AddNewPart<ThemePart>();
                    
                    using (Stream sourceStream = sourceThemePart.GetStream())
                    using (Stream targetStream = targetThemePart.GetStream())
                    {
                        sourceStream.CopyTo(targetStream);
                    }

                    // Set relationships
                    newPresentationPart.Presentation.SlideIdList = new SlideIdList(new SlideId { Id = 256, RelationshipId = newPresentationPart.GetIdOfPart(targetSlidePart) });
                    newPresentationPart.Presentation.Save();
                }
            }
        }

        Console.WriteLine("PowerPoint split completed successfully.");
    }
}
