#include <iostream>
#include <fstream>
#include <string>
#include <vector>
#include <filesystem>
#include <algorithm>
#include "../miniDockReader.h"
#include "../thirdparty/tinyxml2-master/tinyxml2.h"

namespace fs = std::filesystem;

bool isDocxFile(const std::string& filePath) {
    fs::path path(filePath);
    std::string extension = path.extension().string();
    std::transform(extension.begin(), extension.end(), extension.begin(), ::tolower);
    return extension == ".docx";
}

std::string escapeHtmlSpecialChars(const std::string& input) {
    std::string output;
    for (char c : input) {
        switch (c) {
            case '&':
                output += "&amp;";
                break;
            case '<':
                output += "&lt;";
                break;
            case '>':
                output += "&gt;";
                break;
            case '"':
                output += "&quot;";
                break;
            case '\'':
                output += "&#39;";
                break;
            default:
                output += c;
                break;
        }
    }
    return output;
}

bool generateHtmlFromDocument(const Document& doc, const std::string& outputPath) {
    tinyxml2::XMLDocument xmlDoc;

    // Create HTML structure
    tinyxml2::XMLDeclaration* decl = xmlDoc.NewDeclaration("xml version=\"1.0\" encoding=\"UTF-8\"");
    xmlDoc.InsertFirstChild(decl);

    tinyxml2::XMLUnknown* doctype = xmlDoc.NewUnknown("!DOCTYPE html");
    xmlDoc.InsertEndChild(doctype);

    tinyxml2::XMLElement* htmlElement = xmlDoc.NewElement("html");
    xmlDoc.InsertEndChild(htmlElement);

    // Add head
    tinyxml2::XMLElement* headElement = xmlDoc.NewElement("head");
    htmlElement->InsertEndChild(headElement);

    tinyxml2::XMLElement* metaCharset = xmlDoc.NewElement("meta");
    metaCharset->SetAttribute("charset", "UTF-8");
    headElement->InsertEndChild(metaCharset);

    tinyxml2::XMLElement* titleElement = xmlDoc.NewElement("title");
    titleElement->InsertEndChild(xmlDoc.NewText("DOCX Document"));
    headElement->InsertEndChild(titleElement);

    tinyxml2::XMLElement* styleElement = xmlDoc.NewElement("style");
    styleElement->InsertEndChild(xmlDoc.NewText(
        "body { font-family: Arial, sans-serif; margin: 20px; }\n"
        "p { line-height: 1.0; }\n"
        ".run { display: inline; }\n"
        ".bold { font-weight: bold; }\n"
        ".italic { font-style: italic; }\n"
        ".underline { text-decoration: underline; }\n"
        ".strike { text-decoration: line-through; }\n"
        ".subscript { vertical-align: sub; font-size: smaller; }\n"
        ".superscript { vertical-align: super; font-size: smaller; }\n"
        ".heading { margin-top: 20px; margin-bottom: 10px; }\n"
        "h1 { font-size: 28px; }\n"
        "h2 { font-size: 24px; }\n"
        "h3 { font-size: 20px; }"
    ));
    headElement->InsertEndChild(styleElement);

    // Add body
    tinyxml2::XMLElement* bodyElement = xmlDoc.NewElement("body");
    htmlElement->InsertEndChild(bodyElement);

    tinyxml2::XMLElement* titleHeading = xmlDoc.NewElement("h1");
    titleHeading->InsertEndChild(xmlDoc.NewText("Document Contents"));
    bodyElement->InsertEndChild(titleHeading);

    // Add document paragraphs
    if (!doc.paragraphs.empty()) {
        for (size_t i = 0; i < doc.paragraphs.size(); ++i) {
            const Paragraph& para = doc.paragraphs[i];
            
            tinyxml2::XMLElement* pElement = xmlDoc.NewElement("p");
            pElement->SetAttribute("class", "paragraph");

            // Determine and set paragraph styles
            // add embedded styles like alignment, indentation etc. if needed
            std::string paraStyle;
            if (para.alignment == "center") {
                paraStyle += "text-align: center; ";
            } else if (para.alignment == "right") {
                paraStyle += "text-align: right; ";
            } else if (para.alignment == "justify") {
                paraStyle += "text-align: justify; ";
            }
            // add embedded styles like indentation
            if (para.indentLeft > 0) {
                paraStyle += "margin-left: " + std::to_string(para.indentLeft) + "px; ";
            }
            if (para.indentRight > 0) {
                paraStyle += "margin-right: " + std::to_string(para.indentRight) + "px; ";
            }
            if (para.indentFirstLine > 0) {
                paraStyle += "text-indent: " + std::to_string(para.indentFirstLine) + "px; ";
            }
            // add embedded styles like line spacing, space before, and space after
            if (para.lineSpacing != 1.0f) {
                paraStyle += "line-height: " + std::to_string(para.lineSpacing) + "; ";
            }
            if (!para.spaceBefore.empty()) {
                paraStyle += "margin-top: " + std::to_string(para.spaceBefore) + "px; ";
            }
            if (!para.spaceBetweenSameStyle) {
                Paragraph prevPara;
                if (i < doc.paragraphs.size() - 1) {
                    prevPara = doc.paragraphs[i + 1];
                }
                if (&prevPara != nullptr && prevPara.style == para.style) {
                    paraStyle += "margin-bottom: 0px; ";
                }
                else if (!para.spaceAfter.empty()) {
                    paraStyle += "margin-bottom: " + std::to_string(para.spaceAfter) + "px; ";
                }
            }
            else if (!para.spaceAfter.empty()) {
                paraStyle += "margin-bottom: " + std::to_string(para.spaceAfter) + "px; ";
            }
            // direction
            if (para.rightDirection) {
                paraStyle += "direction: rtl; ";
            }

            // Add paragraph styles
            if (!paraStyle.empty()) {
                pElement->SetAttribute("style", paraStyle.c_str());
            }

            // Numbering
            // @todo: add level indentation if needed
            if (para.numbered) {
                tinyxml2::XMLElement* numberElement = xmlDoc.NewElement("span");
                numberElement->SetAttribute("class", "numbering");
                std::string numberingText = para.numberStyle.empty() ? "• " : para.numberStyle + " ";
                numberElement->InsertEndChild(xmlDoc.NewText(numberingText.c_str()));
                pElement->InsertEndChild(numberElement);
            }
            // Tabs
            // @todo: implement tab stops if needed

            // Add runs
            if (para.runs.empty()) {
                // Empty paragraph
                pElement->InsertEndChild(xmlDoc.NewText("\n"));
            } else {
                for (const Run& run : para.runs) {
                    tinyxml2::XMLElement* spanElement = xmlDoc.NewElement("span");
                    spanElement->SetAttribute("class", "run");

                    // add classes for styles
                    std::string runClasses;
                    if (run.bold) runClasses += " bold";
                    if (run.italic) runClasses += " italic";
                    if (run.underline) runClasses += " underline";
                    if (run.strike) runClasses += " strike";
                    if (run.subscript) runClasses += " subscript";
                    if (run.superscript) runClasses += " superscript";
                    if (!runClasses.empty()) {
                        spanElement->SetAttribute("class", runClasses.c_str());
                    }

                    // add embededed styles like color, font-size etc. if needed
                    std::string styleAttr;
                    if (!run.color.empty()) {
                        styleAttr += "color: " + run.color.toHexString() + "; ";
                    }
                    if (!run.backColor.empty()) {
                        styleAttr += "background-color: " + run.backColor.toHexString() + "; ";
                    }
                    if (run.fontFamily.length() > 0) {
                        styleAttr += "font-family: " + run.fontFamily + "; ";
                    }
                    if (run.fontSize > 0) {
                        styleAttr += "font-size: " + std::to_string(run.fontSize) + "pt; ";
                    }
                    if (!styleAttr.empty()) {
                        spanElement->SetAttribute("style", styleAttr.c_str());
                    }

                    // add text content
                    if (run.text.empty()) {
                        spanElement->InsertEndChild(xmlDoc.NewText(""));
                        pElement->InsertEndChild(spanElement);
                    } else {
                        std::string escapedText = escapeHtmlSpecialChars(run.text);
                        spanElement->InsertEndChild(xmlDoc.NewText(escapedText.c_str()));
                        pElement->InsertEndChild(spanElement);
                    }
                }
            }


            bodyElement->InsertEndChild(pElement);
        }
    } else {
        tinyxml2::XMLElement* emptyElement = xmlDoc.NewElement("p");
        emptyElement->InsertEndChild(xmlDoc.NewText("(No content found in document)"));
        bodyElement->InsertEndChild(emptyElement);
    }


    // Save the XML document as HTML
    tinyxml2::XMLError error = xmlDoc.SaveFile(outputPath.c_str());
    if (error != tinyxml2::XML_SUCCESS) {
        std::cerr << "Error saving HTML file: " << error << std::endl;
        return false;
    }

    return true;
}

int main(int argc, char* argv[]) {
    // Check command line arguments
    if (argc < 2) {
        std::cerr << "Usage: " << argv[0] << " <path_to_docx_file>" << std::endl;
        std::cerr << "Example: " << argv[0] << " document.docx" << std::endl;
        return 1;
    }

    std::string filePath = argv[1];

    // Check if file exists
    if (!fs::exists(filePath)) {
        std::cerr << "Error: File does not exist: " << filePath << std::endl;
        return 1;
    }

    // Check if it's a DOCX file
    if (!isDocxFile(filePath)) {
        std::cerr << "Error: File is not a .docx file: " << filePath << std::endl;
        std::cerr << "Please provide a valid DOCX document." << std::endl;
        return 1;
    }

    std::cout << "Processing DOCX file: " << filePath << std::endl;

    // Load the document using miniDockReader library
    Document doc;
    try {
        doc = readDocument(filePath);  // Using your library's function
        std::cout << "Document loaded successfully." << std::endl;
        std::cout << "Found " << doc.paragraphs.size() << " paragraphs." << std::endl;
        std::cout << "Found " << doc.tables.size() << " tables." << std::endl;
    } catch (const std::exception& e) {
        std::cerr << "Error loading DOCX file: " << e.what() << std::endl;
        return 1;
    }

    // Generate output HTML file path
    fs::path inputPath(filePath);
    std::string outputFileName = inputPath.stem().string() + "_output.html";
    fs::path outputPath = inputPath.parent_path() / outputFileName;

    std::cout << "Generating HTML file: " << outputPath.string() << std::endl;

    // Generate HTML from document
    if (!generateHtmlFromDocument(doc, outputPath.string())) {
        std::cerr << "Error: Failed to generate HTML file." << std::endl;
        return 1;
    }

    std::cout << "HTML file generated successfully!" << std::endl;
    std::cout << "Output saved to: " << outputPath.string() << std::endl;

    return 0;
}
