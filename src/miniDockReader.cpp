// miniDockLib_optimized.cpp
// Implementation file for miniDockReader library

#include <string>
#include <vector>
#include <unordered_map>
#include <sstream>
#include <cstring>
#include <cstdlib>
#include <cstdio>

#include "miniDockReader.h"
#include "../thirdparty/miniz-cpp-master/zip_file.hpp"
#include "../thirdparty/tinyxml2-master/tinyxml2.h"

using namespace tinyxml2;

// ---------------- Internal Data Structures ----------------

using StyleMap = std::unordered_map<std::string, Style>;
static std::unordered_map<std::string, Style> g_mergedStyleCache; // cache for merged styles

// -------------- Style merge (cached) --------------
// Merges styles with inheritance, using a cache for performance
// @param styles: map of styleId -> Style
// @param styleId: style ID to merge
// @return merged Style
static Style mergeStyleCached(const StyleMap& styles, const std::string& styleId)
{
    Style result;
    if (styleId.empty()) return result;

    auto itc = g_mergedStyleCache.find(styleId);
    if (itc != g_mergedStyleCache.end())
        return itc->second;

    auto it = styles.find(styleId);
    if (it == styles.end()) {
        g_mergedStyleCache[styleId] = result;
        return result;
    }

    const Style& cur = it->second;
    if (!cur.basedOn.empty())
        result = mergeStyleCached(styles, cur.basedOn);

    // Style type
    // @todo: verify correct behavior here
    if (cur.styleType != ElementType::Paragraph && cur.styleType != ElementType::Run) 
        result.styleType = cur.styleType;

    // Character properties
    if (cur.bold) result.bold = true;
    if (cur.italic) result.italic = true;
    if (cur.underline) result.underline = true;
    if (cur.strikeThrough) result.strikeThrough = true;
    if (cur.Subscript) result.Subscript = true;
    if (cur.Superscript) result.Superscript = true;

    // Colors and font
    if (!cur.color.empty()) result.color = cur.color;
    if (!cur.backColor.empty()) result.backColor = cur.backColor;
    if (!cur.fontFamily.empty()) result.fontFamily = cur.fontFamily;
    if (cur.fontSize > 0) result.fontSize = cur.fontSize;

    // Paragraph properties
    if (cur.lineSpacing > 0) result.lineSpacing = cur.lineSpacing;
    if (cur.spaceBefore > 0) result.spaceBefore = cur.spaceBefore;
    if (cur.spaceAfter > 0) result.spaceAfter = cur.spaceAfter;
    if (cur.spaceBetweenSameStyle) result.spaceBetweenSameStyle = true;
    if (cur.justification != Justification::Left) result.justification = cur.justification;
    if (cur.rightDirection) result.rightDirection = true;
    if (cur.indentLeft > 0) result.indentLeft = cur.indentLeft;
    if (cur.indentRight > 0) result.indentRight = cur.indentRight;
    if (cur.indentFirstLine > 0) result.indentFirstLine = cur.indentFirstLine;

    // Tabs
    if (!cur.tabs.empty()) {
        result.tabs.insert(result.tabs.end(), cur.tabs.begin(), cur.tabs.end());
    }

    // numbering
    if (cur.numbered) result.numbered = true;
    if (!cur.numberFormat.empty()) result.numberFormat = cur.numberFormat;
    if (!cur.numberStyle.empty()) result.numberStyle = cur.numberStyle;
    if (cur.level > 0) result.level = cur.level;

    g_mergedStyleCache[styleId] = result;
    return result;
}

// -------- ZIP: read multiple files --------
// Reads multiple files from a ZIP archive
// @param path: path to ZIP file
// @param files: list of filenames to read
// @return map of filename -> file data
static std::unordered_map<std::string, std::string>
readMultipleFilesFromZIP(const std::string& path,
                         const std::vector<std::string>& files)
{
    std::unordered_map<std::string, std::string> out;
    mz_zip_archive zip;
    std::memset(&zip, 0, sizeof(zip));

    if (!mz_zip_reader_init_file(&zip, path.c_str(), 0))
        return out;

    // Iterate through files in the ZIP
    int n = mz_zip_reader_get_num_files(&zip);
    for (int i = 0; i < n; ++i) {
        mz_zip_archive_file_stat st;
        if (!mz_zip_reader_file_stat(&zip, i, &st)) continue;

        for (const auto& name : files) {
            if (std::strcmp(st.m_filename, name.c_str()) == 0) {
                std::string data;
                data.resize(static_cast<size_t>(st.m_uncomp_size));
                if (!mz_zip_reader_extract_file_to_mem(
                        &zip, st.m_filename, data.data(), st.m_uncomp_size, 0))
                    data.clear();
                out[name] = std::move(data);
                break;
            }
        }
    }

    mz_zip_reader_end(&zip);
    return out;
}

// ---------------- Styles parsing ----------------
// Parses styles.xml and returns a map of styleId -> Style
// @param xml: styles.xml content
// @return map of styleId -> Style
static StyleMap parseStyles(const std::string& xml)
{
    StyleMap map;
    if (xml.empty()) return map;

    XMLDocument doc;
    doc.Parse(xml.c_str());
    XMLElement* root = doc.FirstChildElement("w:styles");
    if (!root) return map;

    // Parse each style
    for (XMLElement* s = root->FirstChildElement("w:style"); s;
         s = s->NextSiblingElement("w:style")) {
        const char* id = s->Attribute("w:styleId");
        if (!id) continue;

        Style st;
        if (const char* t = s->Attribute("w:type")) st.styleType = 
            (std::strcmp(t, "paragraph") == 0) ? ElementType::Paragraph : ElementType::Run;

        // Based on
        if (XMLElement* b = s->FirstChildElement("w:basedOn"))
            if (b->Attribute("w:val")) st.basedOn = b->Attribute("w:val");

        // Run properties
        if (XMLElement* rPr = s->FirstChildElement("w:rPr")) {
            // Character properties
            if (rPr->FirstChildElement("w:b")) st.bold = true;
            if (rPr->FirstChildElement("w:i")) st.italic = true;
            if (rPr->FirstChildElement("w:u")) st.underline = true;
            if (rPr->FirstChildElement("w:strike")) st.strikeThrough = true;
            if (rPr->FirstChildElement("w:subscript")) st.Subscript = true;
            if (rPr->FirstChildElement("w:superscript")) st.Superscript = true;

            // Colors and font
            if (XMLElement* c = rPr->FirstChildElement("w:color"))
                if (c->Attribute("w:val")) st.color = Color(c->Attribute("w:val"));
            if (XMLElement* sh = rPr->FirstChildElement("w:shd"))
                if (sh->Attribute("w:fill")) st.backColor = Color(sh->Attribute("w:fill"));
            if (XMLElement* rf = rPr->FirstChildElement("w:rFonts"))
                if (rf->Attribute("w:ascii")) st.fontFamily = rf->Attribute("w:ascii");
            if (XMLElement* sz = rPr->FirstChildElement("w:sz"))
                if (sz->Attribute("w:val")) st.fontSize = std::stof(sz->Attribute("w:val")) / 2.0f;
        }

        // Paragraph properties
        if (XMLElement* pPr = s->FirstChildElement("w:pPr")) {

            // level
            if (XMLElement* outline = pPr->FirstChildElement("w:outlineLvl"))
                if (outline->Attribute("w:val"))
                    st.level = std::atoi(outline->Attribute("w:val"));
            // Numbering
            if (XMLElement* num = pPr->FirstChildElement("w:numPr")) {
                // Numbering ID
                if (XMLElement* numId = num->FirstChildElement("w:numId"))
                    if (numId->Attribute("w:val"))
                        st.numberFormat = "decimal"; // Default format
                // Level
                if (XMLElement* ilvl = num->FirstChildElement("w:ilvl"))
                    if (ilvl->Attribute("w:val"))
                        st.level = std::atoi(ilvl->Attribute("w:val"));
                // Style
                if (XMLElement* numStyle = num->FirstChildElement("w:numStyle"))
                    if (numStyle->Attribute("w:val"))
                        st.numberStyle = numStyle->Attribute("w:val");
                st.numbered = true;
            }
            // Spacing
            if (XMLElement* sp = pPr->FirstChildElement("w:spacing")) {
                if (sp->Attribute("w:line"))
                    st.lineSpacing = std::stof(sp->Attribute("w:line")) / 240.0f;
                if (sp->Attribute("w:before"))
                    st.spaceBefore = std::stof(sp->Attribute("w:before")) / 20.0f;
                if (sp->Attribute("w:after"))
                    st.spaceAfter = std::stof(sp->Attribute("w:after")) / 20.0f;
                if (sp->Attribute("w:lineRule")) {
                    const char* val = sp->Attribute("w:lineRule");
                    if (std::strcmp(val, "exact") == 0)
                        st.spaceBetweenSameStyle = true;
                }
            }
            // Indentation
            if (XMLElement* indent = pPr->FirstChildElement("w:ind")) {
                if (indent->Attribute("w:left"))
                    st.indentLeft = std::stof(indent->Attribute("w:left")) / 20.0f;
                if (indent->Attribute("w:right"))
                    st.indentRight = std::stof(indent->Attribute("w:right")) / 20.0f;
                if (indent->Attribute("w:firstLine"))
                    st.indentFirstLine = std::stof(indent->Attribute("w:firstLine")) / 20.0f;
            }
            // Justification
            if (XMLElement* jc = pPr->FirstChildElement("w:jc")) {
                if (const char* val = jc->Attribute("w:val")) {
                    if (std::strcmp(val, "center") == 0)
                        st.justification = Justification::Center;
                    else if (std::strcmp(val, "right") == 0)
                        st.justification = Justification::Right;
                    else if (std::strcmp(val, "both") == 0)
                        st.justification = Justification::Justify;
                }
            }
            // Tabs
            if (XMLElement* tabs = pPr->FirstChildElement("w:tabs")) {
                for (XMLElement* tab = tabs->FirstChildElement("w:tab"); tab;
                     tab = tab->NextSiblingElement("w:tab")) {
                    Tab t;
                    if (tab->Attribute("w:pos"))
                        t.position = std::stof(tab->Attribute("w:pos")) / 20.0f;
                    if (tab->Attribute("w:val"))
                        t.alignment = tab->Attribute("w:val")[0]; // L, C, R, D
                    if (tab->Attribute("w:leader"))
                        t.leader = tab->Attribute("w:leader");
                    st.tabs.emplace_back(std::move(t));
                }
            }
            // Bidi
            if (XMLElement* bd = pPr->FirstChildElement("w:bidi"))
                st.rightDirection = true;
        }

        // add the new style to the map
        map.emplace(id, std::move(st));
    }

    return map;
}

// ------------ Text Collector (recursive) ------------
// Recursively collects text from XML elements
// @param e: current XML element
// @param out: output string to collect text
static void collectTextRecursive(XMLElement* e, std::string& out)
{
    if (!e) return;

    if (std::strcmp(e->Name(), "w:t") == 0 && e->GetText()) {
        out += e->GetText();
    }

    for (XMLNode* n = e->FirstChild(); n; n = n->NextSibling()) {
        if (XMLElement* c = n->ToElement()) {
            collectTextRecursive(c, out);
        }
    }
}

// ------------ Parse Footnotes -------------
// Parses footnotes.xml and returns a map of footnote ID -> text
// @param xml: footnotes.xml content
// @return map of footnote ID -> text
// @todo: handle complex footnotes with multiple paragraphs, etc.
// For now, we just collect plain text
// @todo: support footnotes with styles
static std::unordered_map<int, std::string> parseFootnotes(const std::string& xml)
{
    std::unordered_map<int, std::string> map;
    if (xml.empty()) return map;

    XMLDocument doc;
    doc.Parse(xml.c_str());

    XMLElement* root = doc.FirstChildElement("w:footnotes");
    if (!root) return map;

    for (XMLElement* fn = root->FirstChildElement("w:footnote"); fn;
         fn = fn->NextSiblingElement("w:footnote")) {
        const char* idAttr = fn->Attribute("w:id");
        if (!idAttr) continue;
        int id = std::atoi(idAttr);

        const char* typeAttr = fn->Attribute("w:type");
        if (typeAttr &&
            (std::strcmp(typeAttr, "separator") == 0 ||
             std::strcmp(typeAttr, "continuationSeparator") == 0)) {
            continue; // skip separators
        }

        std::string text;
        collectTextRecursive(fn, text);
        map[id] = text;
    }

    return map;
}

// ------------ Compare Run Styles -------------
// Compares two runs to see if they have the same style
// @param a: first Run
// @param b: second Run
// @return true if styles are the same, false otherwise
static bool sameRunStyle(const Run& a, const Run& b)
{
    return a.style == b.style &&
           a.lang == b.lang &&
           a.bold == b.bold &&
           a.italic == b.italic &&
           a.underline == b.underline &&
           a.strike == b.strike &&
           a.color == b.color &&
           a.backColor == b.backColor &&
           a.fontFamily == b.fontFamily &&
           a.fontSize == b.fontSize &&
           a.Subscript == b.Subscript &&
           a.Superscript == b.Superscript;
}

// @todo: check the file from here

// -------- Parse document.xml (with footnotes) --------
// Parses document.xml and returns a vector of Paragraphs
// @param xml: document.xml content
// @param styles: map of styleId -> Style
// @param footnotes: map of footnote ID -> text
// @return vector of Paragraphs
static std::vector<Paragraph> parseDocumentWithFootnotes(
    const std::string& xml,
    const StyleMap&    styles,
    const std::unordered_map<int, std::string>& footnotes)
{
    std::vector<Paragraph> paras;
    if (xml.empty()) return paras;

    XMLDocument doc;
    doc.Parse(xml.c_str());

    XMLElement* root = doc.FirstChildElement("w:document");
    if (!root) return paras;
    XMLElement* body = root->FirstChildElement("w:body");
    if (!body) return paras;

    for (XMLElement* p = body->FirstChildElement("w:p"); p;
         p = p->NextSiblingElement("w:p")) {
        Paragraph  para;
        std::string pStyleId;

        if (XMLElement* pPr = p->FirstChildElement("w:pPr")) {
            if (XMLElement* pStyle = pPr->FirstChildElement("w:pStyle")) {
                if (pStyle->Attribute("w:val"))
                    pStyleId = pStyle->Attribute("w:val");
            }

            if (XMLElement* spacing = pPr->FirstChildElement("w:spacing")) {
                if (spacing->Attribute("w:line"))
                    para.lineSpacing = std::stof(spacing->Attribute("w:line")) / 240.0f;
                if (spacing->Attribute("w:before"))
                    para.spaceBefore = std::stof(spacing->Attribute("w:before")) / 20.0f;
                if (spacing->Attribute("w:after"))
                    para.spaceAfter  = std::stof(spacing->Attribute("w:after"))  / 20.0f;
            }
        }

        Style paraStyle = mergeStyleCached(styles, pStyleId);
        para.style = pStyleId;
        if (paraStyle.lineSpacing > 0.0f) para.lineSpacing = paraStyle.lineSpacing;
        para.spaceBefore = paraStyle.spaceBefore;
        para.spaceAfter  = paraStyle.spaceAfter;

        for (XMLElement* r = p->FirstChildElement("w:r"); r;
             r = r->NextSiblingElement("w:r")) {
            // Footnote reference always creates a new run
            if (XMLElement* fr = r->FirstChildElement("w:footnoteReference")) {
                if (fr->Attribute("w:id")) {
                    int id = std::atoi(fr->Attribute("w:id"));
                    Run run;
                    run.text = "[Footnote " + std::to_string(id) + "]";
                    auto it = footnotes.find(id);
                    if (it != footnotes.end() && !it->second.empty()) {
                        run.text += ": ";
                        run.text += it->second;
                    }
                    para.runs.emplace_back(std::move(run));
                    continue;
                }
            }

            Run run;
            std::string rStyleId;
            if (XMLElement* rPr = r->FirstChildElement("w:rPr")) {
                if (XMLElement* rStyle = rPr->FirstChildElement("w:rStyle")) {
                    if (rStyle->Attribute("w:val"))
                        rStyleId = rStyle->Attribute("w:val");
                }
                if (rPr->FirstChildElement("w:b")) run.bold = true;
                if (rPr->FirstChildElement("w:i")) run.italic = true;
                if (rPr->FirstChildElement("w:u")) run.underline = true;
                if (XMLElement* c = rPr->FirstChildElement("w:color")) {
                    if (c->Attribute("w:val")) run.color = Color(c->Attribute("w:val"));
                }
                if (XMLElement* shd = rPr->FirstChildElement("w:shd")) {
                    if (shd->Attribute("w:fill")) run.backColor = Color(shd->Attribute("w:fill"));
                }
                if (XMLElement* rf = rPr->FirstChildElement("w:rFonts")) {
                    if (rf->Attribute("w:ascii")) run.fontFamily = rf->Attribute("w:ascii");
                }
                if (XMLElement* lang = rPr->FirstChildElement("w:lang")) {
                    if (lang->Attribute("w:val")) run.lang = lang->Attribute("w:val");
                }
                if (XMLElement* sz = rPr->FirstChildElement("w:sz")) {
                    if (sz->Attribute("w:val")) run.fontSize = std::stof(sz->Attribute("w:val")) / 2.0f;
                }
            }

            Style rStyle = mergeStyleCached(styles, rStyleId);
            if (rStyle.bold) run.bold = true;
            if (rStyle.italic) run.italic = true;
            if (rStyle.underline) run.underline = true;
            if (!rStyle.color.empty()) run.color = rStyle.color;
            if (!rStyle.backColor.empty()) run.backColor = rStyle.backColor;
            if (!rStyle.fontFamily.empty()) run.fontFamily = rStyle.fontFamily;
            if (rStyle.fontSize > 0.0f) run.fontSize = rStyle.fontSize;
            run.style = rStyleId;

            if (XMLElement* t = r->FirstChildElement("w:t")) {
                if (t->GetText()) run.text = t->GetText();
            } else {
                std::string tmp;
                collectTextRecursive(r, tmp);
                run.text = tmp;
            }

            if (!para.runs.empty() && sameRunStyle(para.runs.back(), run)) {
                para.runs.back().text += run.text;
            } else {
                para.runs.emplace_back(std::move(run));
            }
        }

        paras.emplace_back(std::move(para));
    }

    return paras;
}

// ---------------- Public API Functions ----------------
// Read document from file path
MINIDOCKLIB_API std::vector<Paragraph> readDocument(
    const std::string& path)
{
    g_mergedStyleCache.clear();

    std::vector<std::string> filesToRead = {
        "word/document.xml",
        "word/styles.xml",
        "word/footnotes.xml"
    };

    auto fileData = readMultipleFilesFromZIP(path, filesToRead);

    StyleMap styles = parseStyles(fileData["word/styles.xml"]);
    auto footnotes = parseFootnotes(fileData["word/footnotes.xml"]);
    auto paragraphs = parseDocumentWithFootnotes(
        fileData["word/document.xml"], styles, footnotes);

    return paragraphs;
}

// Read document from memory buffer
MINIDOCKLIB_API std::vector<Paragraph> readDocumentFromMemory(
    const char* data,
    size_t      size)
{
    g_mergedStyleCache.clear();

    std::vector<std::string> filesToRead = {
        "word/document.xml",
        "word/styles.xml",
        "word/footnotes.xml"
    };

    // Create a temporary file to hold the in-memory data
    char tempPath[L_tmpnam];
    std::tmpnam(tempPath);
    FILE* tempFile = std::fopen(tempPath, "wb");
    if (!tempFile) return {};

    std::fwrite(data, 1, size, tempFile);
    std::fclose(tempFile);

    auto fileData = readMultipleFilesFromZIP(tempPath, filesToRead);

    // Remove the temporary file
    std::remove(tempPath);

    StyleMap styles = parseStyles(fileData["word/styles.xml"]);
    auto footnotes = parseFootnotes(fileData["word/footnotes.xml"]);
    auto paragraphs = parseDocumentWithFootnotes(
        fileData["word/document.xml"], styles, footnotes);

    return paragraphs;
}

// TODO: continue from here with other functions exposed in miniDockReader.h