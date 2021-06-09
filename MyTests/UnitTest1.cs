using System;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Xunit;

namespace MyTests
{
    public class UnitTest1
    {

            [Fact]
            public void OXS875Testing()
            {
                using (var presentationDocument = PresentationDocument.Open(@"c:\Tasks\OXS875\OXS-875.pptx", false))
                {
                    var presentationPart = presentationDocument.PresentationPart;
                    var presentation = presentationPart.Presentation;

                    var slideIdList = presentation.SlideIdList;

                    foreach (var slideId in slideIdList.ChildElements.OfType<SlideId>())
                    {
                        SlidePart slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId);

                        var slide = slidePart.Slide;

                        var background = slide.CommonSlideData.Background;
                        var backgroundProperties = background.BackgroundProperties;
                        var solidFill = backgroundProperties.GetFirstChild<SolidFill>();
                        var solidFillRgbColorModelHex = solidFill.RgbColorModelHex;
                        var alpha = solidFillRgbColorModelHex.GetFirstChild<Alpha>();
                        try
                        {
                            int alphaVal = alpha.Val;
                        }
                        catch (Exception e)
                        {
                            // Input string was not in a correct format.
                            Console.WriteLine(e);
                        }
                    }
                }
            }
    }
}
