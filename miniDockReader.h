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
// * GitHub: https://github.com/tfussell/miniz-cpp
// * License: MIT License
// tinyxml2:
// * GitHub: https://github.com/leethomason/tinyxml2
// * License: zlib License


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
enum class ElementType {
    Paragraph,
    Run
};

enum class Justification {
    Left,
    Center,
    Right,
    Justify
};

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

struct Tab {
    float position = 0.0f; // in points
    char  alignment = 'L'; // L, C, R, D (decimal)
    std::string leader; // e.g. dots
};
struct Style {
    ElementType styleType;     // paragraph or run
    std::string basedOn;
    bool        bold        = false;
    bool        italic      = false;
    bool        underline   = false;
    bool        strikeThrough = false;
    bool        Subscript     = false;
    bool        Superscript   = false;
    Color       color;
    Color       backColor;
    std::string fontFamily;
    float       fontSize    = 0.0f;

    // paragraph properties
    // numbering
    int         level = 0;
    bool        numbered = false;
    std::string  numberFormat;  // e.g. decimal, upperRoman, lowerLetter
    std::string  numberStyle;   // e.g. "1.", "(a)", etc.
    // spacing
    float       lineSpacing = 1.0f;
    float       spaceBefore = 0.0f;
    float       spaceAfter  = 0.0f;
    bool        spaceBetweenSameStyle = false;
    // alignment
    Justification justification = Justification::Left;
    bool        rightDirection = false;
    // indentation
    float       indentLeft = 0.0f;
    float       indentRight = 0.0f;
    float       indentFirstLine = 0.0f;
    // tabs
    std::vector<Tab> tabs;
};

struct Paragraph {
    std::string style;  // style ID

    // numbering
    int         level = 0;
    bool        numbered = false;
    std::string  numberFormat;  // e.g. decimal, upperRoman, lowerLetter
    std::string  numberStyle;   // e.g. "1.", "(a)", etc.
    // alignment
    Justification justification = Justification::Left;
    bool        rightDirection = false;
    // spacing
    float            lineSpacing = 1.0f;
    float            spaceBefore = 0.0f;
    float            spaceAfter  = 0.0f;
    bool             spaceBetweenSameStyle = false;
    // indentation
    float       indentLeft = 0.0f;
    float       indentRight = 0.0f;
    float       indentFirstLine = 0.0f;
    // tabs
    std::vector<Tab> tabs;

    // runs
    std::vector<Run> runs;
};

struct Run {
    std::string text;
    std::string lang;
    std::string style;  // style ID
    bool        bold        = false;
    bool        italic      = false;
    bool        underline   = false;
    bool        strike      = false;
    bool        Subscript   = false;
    bool        Superscript = false;
    Color       color;
    Color       backColor;
    std::string fontFamily;
    float       fontSize    = 0.0f; // points
};


using StyleMap = std::unordered_map<std::string, Style>;

// API function declarations

// Reads a MiniDock document from a file path
// @param path: path to the MiniDock (.docx) file
// @return vector of Paragraphs
MINIDOCKLIB_API std::vector<Paragraph> readDocument(
      const std::string& path);

// Reads a MiniDock document from in-memory data
// @param data: pointer to the in-memory data
// @param size: size of the in-memory data
// @return vector of Paragraphs
MINIDOCKLIB_API std::vector<Paragraph> readDocumentFromMemory(
    const char* data,
    size_t      size);