#pragma once

// --- Visual representation of the result matrix ---
constexpr int LINESIZE = 8;   // thickness of thick border lines
constexpr int OFFSET_Y = 100; // height of the column-name header row
constexpr int OFFSET_X = 100; // width of the row-name header column
constexpr int STEP_X = 24;    // cell width in pixels
constexpr int STEP_Y = 24;    // cell height in pixels

// --- Algorithm limits ---
constexpr int SUGKEYS = 10;
constexpr int MAX_ATTEMPTS = 1000000;

// --- File dialog ---
constexpr int MAX_CFileDialog_FILE_COUNT = 1;
constexpr int FILE_LIST_BUFFER_SIZE = (MAX_CFileDialog_FILE_COUNT * (MAX_PATH + 1)) + 1;

// --- Custom window messages: progress (lParam = 0–100) ---
constexpr UINT CM_UPDATE_PROGRESS = WM_APP + 1;
constexpr UINT CM_UPDATE_PROGRESS2 = WM_APP + 2;
constexpr UINT CM_UPDATE_PROGRESS3 = WM_APP + 3;
constexpr UINT CM_UPDATE_KEYPROGRESS1 = WM_APP + 4;
constexpr UINT CM_UPDATE_KEYPROGRESS2 = WM_APP + 5;

// --- Custom window messages: worker-thread completion events ---
constexpr UINT CM_KEYS1_DONE = WM_APP + 6;      // createKeyArrays1 complete
constexpr UINT CM_KEYS2_DONE = WM_APP + 7;      // createKeyArrays2 complete
constexpr UINT CM_GATHERING1_DONE = WM_APP + 8; // suggestKeys1 complete
constexpr UINT CM_GATHERING2_DONE = WM_APP + 9; // suggestKeys2 complete
constexpr UINT CM_KEYS_FOUND = WM_APP + 10;     // mutualCheck succeeded
constexpr UINT CM_KEYS_NOT_FOUND = WM_APP + 11; // mutualCheck failed
constexpr UINT CM_MARKING_READY = WM_APP + 12;  // markInFiles done, trigger post-processing
constexpr UINT CM_SIMS1_DONE = WM_APP + 13;     // findSims1 complete
constexpr UINT CM_SIMS2_DONE = WM_APP + 14;     // findSims2 complete
constexpr UINT CM_FIRSTPASS_DONE = WM_APP + 15; // firstPass complete

// --- Utility functions (replaced function-like macros) ---
template <typename T> constexpr int sgn(T x)
{
    return static_cast<int>((x > T(0)) - (x < T(0)));
}

// --- Helpers for using MFC CString as unordered_map key ---
#include <string>
#include <unordered_map>

/// @brief Hasher for MFC CString — delegates to std::hash<std::wstring>.
struct CStringHash
{
    std::size_t operator()(const CString& s) const noexcept
    {
        return std::hash<std::wstring>{}(std::wstring(s.GetString(), s.GetLength()));
    }
};
/// @brief Equality comparator for MFC CString (case-sensitive).
struct CStringEqual
{
    bool operator()(const CString& a, const CString& b) const noexcept { return a == b; }
};
