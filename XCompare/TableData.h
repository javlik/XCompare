#pragma once

/** @brief Up to 256 candidate key column indices for one table, as selected by the key-suggestion algorithm. */
struct PossibleKeys
{
    int k[256]; ///< Column indices (1-based) of candidate key columns.
};

/** @brief One RGB entry in the colour palette used to highlight differences. */
struct Palette
{
    int red;   ///< Red channel (0–255).
    int green; ///< Green channel (0–255).
    int blue;  ///< Blue channel (0–255).
};

/** @brief Describes the layout of one loaded Excel sheet (table). */
struct Table
{
    int WorkSheetNumber;  ///< 1-based sheet index within the workbook.
    long MaxNumberOfRows; ///< Maximum row capacity (used-range upper bound from Excel).
    long MaxNumberOfCols; ///< Maximum column capacity (used-range upper bound from Excel).
    long NumberOfRows;    ///< Actual number of data rows (excluding the header row).
    int FirstRowWithData; ///< 1-based row index of the first data row.
    int RowWithNames;     ///< 1-based row index of the column-name header row.
    int NumberOfColumns;  ///< Number of columns present in the sheet.
    CString Columns[256]; ///< Column header names, indexed 1–NumberOfColumns.
    bool keys[256];       ///< Per-column flag: true if that column is part of the active key.
    int keysCnt;          ///< Number of key columns currently selected.
};

/** @brief Top-left corner of the visible comparison matrix, in cell-unit coordinates. */
struct VisTopLeft
{
    int top;  ///< First visible row index (0-based).
    int left; ///< First visible column index (0-based).
};

/** @brief A pair of matched key columns, one from each table. */
struct KeyPair
{
    int tab1; ///< Key column index in table 1 (1-based).
    int tab2; ///< Key column index in table 2 (1-based).
};

/** @brief Pairwise similarity score between one column from each table. */
struct SimilaritiesAcrossTables
{
    int clm1;            ///< Column index in table 1 (1-based).
    int clm2;            ///< Column index in table 2 (1-based).
    long similarity;     ///< Raw overlap count (number of matching values).
    int similarityOrder; ///< Rank of this pair in the sorted similarity list.
    int pureSim;         ///< Similarity after removing exact-name-match bonus.
};

/** @brief Screen coordinates (in cell units) of the cell the user has clicked on. */
struct ChosenCell
{
    int x; ///< Column index (1-based, 0 = none).
    int y; ///< Row index (1-based, 0 = none).
};

/** @brief Pixel dimensions of the client drawing area. */
struct Clnt
{
    int w; ///< Width in pixels.
    int h; ///< Height in pixels.
};

/** @brief Best key-column combination found so far by the key-suggestion algorithm. */
struct BestKeyComb
{
    int pk1;    ///< Index into PossibleKeys for table 1.
    int pk2;    ///< Index into PossibleKeys for table 2.
    int rating; ///< Combined score (higher is better).
    long cnt;   ///< Number of successfully matched rows with this combination.
};

/** @brief Details of the first duplicate key value encountered during uniqueness checking. */
struct NotUniqueKeys
{
    long firstRow;     ///< Row of the first occurrence of the duplicate.
    long secondRow;    ///< Row of the second occurrence of the duplicate.
    CString keyString; ///< The duplicated key value (as a string).
};
