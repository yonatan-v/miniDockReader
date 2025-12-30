#pragma once

// miniDockLib.h
// Public API for miniDockReader library

// Copyright (c) 2024 MiniDockReader Developers
// The project homepage in GitHub: https://github.com/yonatan-v/miniDockReader
// Licensed under the MIT License
// See LICENSE file in the project root for full license information.
//
// THIRD PARTY LICENSES:
// thirdparty libraries used:
// miniz-cpp:
//   GitHub: https://github.com/tfussell/miniz-cpp
//   License: MIT License
// tinyxml2:
//   GitHub: https://github.com/leethomason/tinyxml2
//   License: zlib License


// ---------------- Description ----------------
// Usage:
// This library provides functions to read and parse MiniDock (.docx) documents.
// It exposes functions to read documents from file paths or from in-memory data,
// returning structured representations of the document content.
//
// Example:
//   std::vector<Paragraph> paragraphs = readDocument("example.docx");
//   std::vector<Paragraph> paragraphsFromMemory = readDocumentFromMemory(data, size);
// ---------------- End Description ----------------

// Project uses C++17 standard
#include <cstdint>
#include <unordered_map>
#include <vector>
#include <string>

// ---- Compiler / platform helpers ----
#if defined(_MSC_VER)
  #define MINIDOCKLIB_COMPILER_MSVC 1
#elif defined(__clang__)
  #define MINIDOCKLIB_COMPILER_CLANG 1
#elif defined(__GNUC__)
  #define MINIDOCKLIB_COMPILER_GCC 1
#else
  #define MINIDOCKLIB_COMPILER_UNKNOWN 1
#endif

#if defined(_WIN32) || defined(_WIN64)
  #define MINIDOCKLIB_PLATFORM_WINDOWS 1
#else
  #define MINIDOCKLIB_PLATFORM_POSIX 1
#endif

// ---- DLL import / export (optional future use) ----
#if defined(MINIDOCKLIB_BUILD_DLL)
  #if defined(MINIDOCKLIB_PLATFORM_WINDOWS)
    #if defined(MINIDOCKLIB_EXPORTS)
      #define MINIDOCKLIB_API __declspec(dllexport)
    #else
      #define MINIDOCKLIB_API __declspec(dllimport)
    #endif
  #else
    #define MINIDOCKLIB_API __attribute__((visibility("default")))
  #endif
#else
  #define MINIDOCKLIB_API
#endif

// ---- Public API ----
// Element type enum
enum class ElementType {
    Paragraph,
    Run
};

// Justification enum
enum class Justification {
    Left,
    Center,
    Right,
    Justify
};

// Color structure
// Represents RGBA color
struct Color {
    uint8_t r = 0;
    uint8_t g = 0;
    uint8_t b = 0;
    uint8_t a = 255;

    Color() = default;
    Color(uint8_t red, uint8_t green, uint8_t blue, uint8_t alpha = 255)
        : r(red), g(green), b(blue), a(alpha) {}
    Color(const std::string& hex) {
        if (hex.length() == 6) {
            r = static_cast<uint8_t>(std::stoi(hex.substr(0, 2), nullptr, 16));
            g = static_cast<uint8_t>(std::stoi(hex.substr(2, 2), nullptr, 16));
            b = static_cast<uint8_t>(std::stoi(hex.substr(4, 2), nullptr, 16));
        }
        else if (hex.length() == 8) {
            r = static_cast<uint8_t>(std::stoi(hex.substr(0, 2), nullptr, 16));
            g = static_cast<uint8_t>(std::stoi(hex.substr(2, 2), nullptr, 16));
            b = static_cast<uint8_t>(std::stoi(hex.substr(4, 2), nullptr, 16));
            a = static_cast<uint8_t>(std::stoi(hex.substr(6, 2), nullptr, 16));
        }
    }

    bool empty() const {
        return r == 0 && g == 0 && b == 0 && a == 255;
    }

    bool operator==(const Color& other) const {
        return r == other.r && g == other.g && b == other.b && a == other.a;
    }
};

// Tab stop structure
struct Tab {
    float position = 0.0f;              // in points
    char  alignment = 'L';              // L, C, R, D (decimal)
    std::string leader;                 // e.g. dots
};

// Style structure representing both paragraph and run styles
// Used for merging styles
struct Style {
    ElementType styleType;              // paragraph or run
    std::string basedOn;                // the style which this style based on
    bool        bold        = false;    // the 'bold' style
    bool        italic      = false;    // the 'italic' style
    bool        underline   = false;    // the text has line under the text
    bool        strikeThrough = false;  // the text has line through the text
    bool        Subscript     = false;  // the text is subscript, e.g. for footnotes
    bool        Superscript   = false;  // the text is superscript
    Color       color;                  // the color of the text
    Color       backColor;              // the background color of the text
    std::string fontFamily;             // font family name (e.g. Arial)
    float       fontSize    = 0.0f;     // font size in points

    // paragraph properties
    // numbering
    int         level = 0;              // the numbering level
    bool        numbered = false;       // is numbered paragraph
    std::string  numberFormat;          // e.g. decimal, upperRoman, lowerLetter
    std::string  numberStyle;           // e.g. "1.", "(a)", etc.
    // spacing
    float       lineSpacing = 1.0f;     // line spacing multiplier
    float       spaceBefore = 0.0f;     // space before the paragraph
    float       spaceAfter  = 0.0f;     // space after the paragraph
    bool        spaceBetweenSameStyle = false;  // special handling for same-style paragraphs
    // alignment
    Justification justification = Justification::Left;  // paragraph alignment
    bool        rightDirection = false; // is right-to-left direction
    // indentation
    float       indentLeft = 0.0f;      // left indent in points
    float       indentRight = 0.0f;     // right indent in points
    float       indentFirstLine = 0.0f; // first line indent in points
    // tabs
    std::vector<Tab> tabs;              // tab stops
};

// Paragraph structure
// Represents a paragraph in the document
struct Paragraph {
    std::string style;                  // style ID

    // numbering
    int         level = 0;              // the numbering level
    bool        numbered = false;       // is numbered paragraph
    std::string numberFormat;           // e.g. decimal, upperRoman, lowerLetter
    std::string numberStyle;            // e.g. "1.", "(a)", etc.
    // alignment
    Justification justification = Justification::Left;  // paragraph alignment
    bool        rightDirection = false; // is right-to-left direction
    // spacing
    float       lineSpacing = 1.0f;     // line spacing multiplier
    float       spaceBefore = 0.0f;     // space before the paragraph
    float       spaceAfter  = 0.0f;     // space after the paragraph
    bool        spaceBetweenSameStyle = false; // special handling for same-style paragraphs
    // indentation
    float       indentLeft = 0.0f;      // left indent in points
    float       indentRight = 0.0f;     // right indent in points
    float       indentFirstLine = 0.0f; // first line indent in points
    // tabs
    std::vector<Tab> tabs;              // tab stops

    // runs
    std::vector<Run> runs;              // vector of runs in the paragraph
};

// Run structure
// Represents a text run in a paragraph
struct Run {
    std::string text;                   // run text
    std::string lang;                   // style properties
    std::string style;                  // style ID
    bool        bold        = false;    // the 'bold' style
    bool        italic      = false;    // the 'italic' style
    bool        underline   = false;    // the text has line under the text
    bool        strike      = false;    // the text has line through the text
    bool        subscript   = false;    // the text is subscript, e.g. for footnotes
    bool        superscript = false;    // the text is superscript
    Color       color;                  // the color of the text
    Color       backColor;              // the background color of the text     
    std::string fontFamily;             // font family name (e.g. Arial)
    float       fontSize    = 0.0f;     // font size in points

    // for Notes:
    uint32_t    noteId      = 0;        // the footnote / endnote ID
};

// Note structure
// Represents a footnote or endnote
struct Note{
    int         id;                     // footnote ID
    std::vector<Paragraph> paragraphs;  // footnote text
};

// Document structure
// Represents the entire document
struct Document {
    std::vector<Paragraph> paragraphs;  // vector of paragraphs in the document
    std::unordered_map<std::string, Style> styles; // map of style ID to Style
    std::unordered_map<int, Note> footnotes; // map of footnote ID to Note
    std::unordered_map<int, Note> endnotes;  // map of endnote ID to Note
};



// API function declarations

// Reads a MiniDock document from a file path
// @param path: path to the MiniDock (.docx) file
// @return Document structure representing the document
MINIDOCKLIB_API Document readDocument(
      const std::string& path);

// Reads a MiniDock document from in-memory data
// @param data: pointer to the in-memory data
// @param size: size of the in-memory data
// @return Document structure representing the document
MINIDOCKLIB_API Document readDocumentFromMemory(
    const char* data,
    size_t      size);
