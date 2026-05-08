//
//  Copyright (C) 2005 Serge Wautier - appTranslator
//
//  appTranslator - The ultimate localization tool for your Visual C++ applications
//                  http://www.apptranslator.com
//
//  This source code is provided 'as-is', without any express or implied
//  warranty. In no event will the author be held liable for any damages
//  arising from the use of this software.
//
//  Permission is granted to anyone to use this software for any purpose,
//  including commercial applications, and to alter it and redistribute it
//  freely, subject to the following restrictions:
//
//  1. The origin of this source code must not be misrepresented; you must not
//    claim that you wrote the original source code. If you use this source code
//    in a product, an acknowledgment in the product documentation and in the
//    About box is required, mentioning appTranslator and http://www.apptranslator.com
//
//  2. Altered source versions must be plainly marked as such, and must not be
//    misrepresented as being the original source code.
//
//  3. This notice may not be removed or altered from any source distribution
//

//
// Msg.h: Declaration of the CMsg and CFMsg classes
//

#pragma once

/**
 * @brief CString loaded from the resource string table by ID.
 *
 * Pass a string-table resource ID to the constructor; the resulting object
 * can be used wherever a @c CString or @c LPCTSTR is expected.
 *
 * @code
 *   AfxMessageBox(CMsg(IDS_ERROR), MB_ICONERROR);
 * @endcode
 */
class CMsg : public CString
{
public:
    /** @brief Loads string @p nID from the resource string table. Asserts in debug if not found. */
    CMsg(UINT nID);
};

/**
 * @brief Formatted message string built with FormatMessage-style arguments.
 *
 * A super-printf that uses @c FormatMessage() notation (@c %1, @c %2!d!, …)
 * instead of @c printf notation. The result is a @c CString and can be used
 * wherever a @c LPCTSTR is expected.
 *
 * @code
 *   // IDS_AGE resource: "%1 is %2!d! years old."
 *   AfxMessageBox(CFMsg(IDS_AGE, szName, nAge), MB_ICONINFORMATION);
 * @endcode
 */
class CFMsg : public CString
{
public:
    /** @brief Builds a formatted string from a literal format string (FormatMessage notation). */
    CFMsg(LPCTSTR pszFormat, ...);
    /** @brief Builds a formatted string from a string-table format resource. */
    CFMsg(UINT nFormatID, ...);
};
