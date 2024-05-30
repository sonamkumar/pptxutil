using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace PowerPointSplitter
{
    class Program
    {
        static void Main(string[] args)
        {
            string sourcePresentationPath = "path/to/your/presentation.pptx";
            string outputDirectory = "path/to/output/directory";

            using (PresentationDocument sourceDoc = PresentationDocument.Open(sourcePresentationPath, false))
            {
                PresentationPart sourcePresentationPart = sourceDoc.PresentationPart;
                SlideIdList slideIdList = sourcePresentationPart.Presentation.SlideIdList;

                foreach (SlideId slideId in slideIdList.ChildElements)
                {
                    string slideName = ((SlidePart)sourcePresentationPart.GetPartById(slideId.RelationshipId)).Slide.Name;
                    string newPresentationPath = Path.Combine(outputDirectory, $"{slideName}.pptx");

                    using (PresentationDocument newDoc = PresentationDocument.Create(newPresentationPath, PresentationDocumentType.Presentation))
                    {
                        PresentationPart newPresentationPart = newDoc.AddPresentationPart();
                        newPresentationPart.Presentation = new Presentation();

                        // Create new SlideIdList and add the current slide
                        SlideIdList newSlideIdList = new SlideIdList();
                        newSlideIdList.Append(new SlideId() { Id = slideId.Id, RelationshipId = slideId.RelationshipId });
                        newPresentationPart.Presentation.Append(newSlideIdList);

                        // Add SlideParts
                        SlidePart sourceSlidePart = (SlidePart)sourcePresentationPart.GetPartById(slideId.RelationshipId);
                        SlidePart newSlidePart = newPresentationPart.AddPart<SlidePart>(sourceSlidePart);

                        // Add SlideMasterPart if exists
                        SlideMasterPart sourceSlideMasterPart = sourceSlidePart.SlideLayoutPart.SlideMasterPart;
                        if (sourceSlideMasterPart != null)
                        {
                            SlideMasterPart newSlideMasterPart = newPresentationPart.AddPart<SlideMasterPart>(sourceSlideMasterPart);

                            // Update SlideLayoutIds in SlideMasterPart
                            foreach (SlideLayoutId layoutId in newSlideMasterPart.SlideMaster.SlideLayoutIdList.ChildElements)
                            {
                                layoutId.RelationshipId = newPresentationPart.GetIdOfPart(newSlideMasterPart.GetPartById(layoutId.RelationshipId));
                            }
                        }

                        // Update SlideLayoutId in SlidePart
                        newSlidePart.SlideLayoutPart.SlideLayoutId.RelationshipId = newPresentationPart.GetIdOfPart(newSlidePart.SlideLayoutPart.SlideMasterPart);
                    }
                }
            }
        }
    }
}
