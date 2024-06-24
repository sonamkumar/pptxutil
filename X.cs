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
                    var newPresentationPart = newPresentation.AddPresentationPart();
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

                    // Copy slide layout and its relationships
                    var sourceSlideLayoutPart = sourceSlidePart.SlideLayoutPart;
                    var targetSlideLayoutPart = targetSlidePart.AddNewPart<SlideLayoutPart>();
                    CopyPart(sourceSlideLayoutPart, targetSlideLayoutPart);

                    // Copy slide master and its relationships
                    var sourceSlideMasterPart = sourceSlideLayoutPart.SlideMasterPart;
                    var targetSlideMasterPart = targetSlideLayoutPart.AddNewPart<SlideMasterPart>();
                    CopyPart(sourceSlideMasterPart, targetSlideMasterPart);

                    // Copy theme
                    var sourceThemePart = sourceSlideMasterPart.ThemePart;
                    var targetThemePart = targetSlideMasterPart.AddNewPart<ThemePart>();
                    CopyPart(sourceThemePart, targetThemePart);

                    // Copy any other required parts (e.g., image parts)
                    CopyReferencedParts(sourceSlidePart, targetSlidePart);

                    // Set relationships
                    newPresentationPart.Presentation.SlideIdList = new SlideIdList(new SlideId { Id = 256, RelationshipId = newPresentationPart.GetIdOfPart(targetSlidePart) });
                    newPresentationPart.Presentation.Save();
                }
            }
        }

        Console.WriteLine("PowerPoint split completed successfully.");
    }

    static void CopyPart(OpenXmlPart sourcePart, OpenXmlPart targetPart)
    {
        using (Stream sourceStream = sourcePart.GetStream())
        using (Stream targetStream = targetPart.GetStream())
        {
            sourceStream.CopyTo(targetStream);
        }

        foreach (var idPair in sourcePart.Parts)
        {
            var sourcePart1 = idPair.OpenXmlPart;
            var targetPart1 = targetPart.AddNewPart(sourcePart1.ContentType, idPair.RelationshipId);
            CopyPart(sourcePart1, targetPart1);
        }
    }

    static void CopyReferencedParts(SlidePart sourceSlidePart, SlidePart targetSlidePart)
    {
        var imageParts = sourceSlidePart.ImageParts;
        foreach (var imagePart in imageParts)
        {
            var targetImagePart = targetSlidePart.AddImagePart(imagePart.ContentType);
            using (Stream sourceStream = imagePart.GetStream())
            using (Stream targetStream = targetImagePart.GetStream())
            {
                sourceStream.CopyTo(targetStream);
            }
        }

        // Copy other types of parts as needed (e.g., EmbeddedObjectParts, AudioParts, VideoParts)
    }
}
