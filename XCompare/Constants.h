#pragma once

// --- Visual representation of the result matrix ---
constexpr int LINESIZE = 8;   // thickness of thick border lines
constexpr int OFFSET_Y = 100; // height of the column-name header row
constexpr int OFFSET_X = 100; // width of the row-name header column
constexpr int STEP_X   = 24;  // cell width in pixels
constexpr int STEP_Y   = 24;  // cell height in pixels

// --- Algorithm limits ---
constexpr int SUGKEYS      = 10;
constexpr int MAX_ATTEMPTS = 1000000;

// --- File dialog ---
constexpr int MAX_CFileDialog_FILE_COUNT = 1;
constexpr int FILE_LIST_BUFFER_SIZE = (MAX_CFileDialog_FILE_COUNT * (MAX_PATH + 1)) + 1;

// --- Custom window messages ---
constexpr UINT CM_UPDATE_PROGRESS     = WM_APP + 1;
constexpr UINT CM_UPDATE_PROGRESS2    = WM_APP + 2;
constexpr UINT CM_UPDATE_PROGRESS3    = WM_APP + 3;
constexpr UINT CM_UPDATE_KEYPROGRESS1 = WM_APP + 4;
constexpr UINT CM_UPDATE_KEYPROGRESS2 = WM_APP + 5;

// --- Utility functions (replaced function-like macros) ---
template<typename T>
constexpr int sgn(T x) { return static_cast<int>((x > T(0)) - (x < T(0))); }
