#pragma once
#include "TableData.h"
#include "ComparisonMatrix.h"
#include "Constants.h"
#include "ExcelConnector.h"
#include <vector>
#include <map>

/**
 * @brief Encapsulates the comparison algorithm state and logic for both tables.
 *
 * Holds all per-comparison data arrays, key structures and the lookup maps.
 * Requires a @c HWND for progress @c PostMessage calls and references to the
 * two @c ExcelConnector objects so it can read cell values.
 *
 * Typical usage:
 * @code
 *   ComparisonEngine m_engine;
 *   m_engine.init(hWnd, m_excel1, m_excel2, m_Table1, m_Table2);
 *   m_engine.makePrereq1();
 *   m_engine.firstPass(...);
 * @endcode
 */
class ComparisonEngine
{
public:
    // --- Initialisation ---
    /** @brief Initialises the engine with window handle, Excel connections and table descriptors. */
    void init(HWND hWnd,
              ExcelConnector& excel1, ExcelConnector& excel2,
              const Table& table1, const Table& table2)
    {
        m_hWnd   = hWnd;
        m_pExcel1 = &excel1;
        m_pExcel2 = &excel2;
        m_Table1 = table1;
        m_Table2 = table2;
    }

    /** @brief Refreshes the table descriptors after a sheet selection change. */
    void setTables(const Table& table1, const Table& table2)
    {
        m_Table1 = table1;
        m_Table2 = table2;
    }

    // --- Prerequisite building ---
    /** @brief Reads all cells of table 1 into internal char arrays and marks empty columns. */
    void makePrereq1()
    {
        m_bPrereq1valid = false;
        m_pchMainArr1.assign((m_Table1.NumberOfColumns + 1) * (m_Table1.NumberOfRows + 1), 0);
        makeCharArr1();
        checkEmptiness1();
        m_bPrereq1valid = true;
    }

    /** @brief Reads all cells of table 2 into internal char arrays and marks empty columns. */
    void makePrereq2()
    {
        m_bPrereq2valid = false;
        m_pchMainArr2.assign((m_Table2.NumberOfColumns + 1) * (m_Table2.NumberOfRows + 1), 0);
        makeCharArr2();
        checkEmptiness2();
        m_bPrereq2valid = true;
    }

    /** @brief Returns @c true if the prerequisite data for table 1 is up to date. */
    bool isPrereq1Valid() const { return m_bPrereq1valid; }
    /** @brief Returns @c true if the prerequisite data for table 2 is up to date. */
    bool isPrereq2Valid() const { return m_bPrereq2valid; }
    /** @brief Marks the table-1 prerequisite data as stale (e.g. after a sheet change). */
    void invalidatePrereq1()    { m_bPrereq1valid = false; }
    /** @brief Marks the table-2 prerequisite data as stale (e.g. after a sheet change). */
    void invalidatePrereq2()    { m_bPrereq2valid = false; }

    // --- Key arrays ---
    /**
     * @brief Builds the concatenated key string array for table 1 and the lookup map.
     * @return 0 on success, 1 if duplicate keys were found in table 1.
     */
    int createKeyArrays1() { return createKeyArraysImpl(1); }
    /**
     * @brief Builds the concatenated key string array for table 2 and the lookup map.
     * @return 0 on success, 2 if duplicate keys were found in table 2.
     */
    int createKeyArrays2() { return createKeyArraysImpl(2); }

    /** @brief Verifies that every key string in table 1 is unique. @return @c true if all keys are unique. */
    bool checkKeysUniqueness1() { return checkKeysUniquenessImpl(1); }
    /** @brief Verifies that every key string in table 2 is unique. @return @c true if all keys are unique. */
    bool checkKeysUniqueness2() { return checkKeysUniquenessImpl(2); }

    // --- First pass (main comparison algorithm) ---
    /**
     * @brief Runs the main comparison pass and fills the similarity matrix.
     *
     * Matches each row of table 1 to the corresponding row in table 2 via the key
     * lookup map, then counts matching cell values per column pair.
     * Posts @c CM_FIRSTPASS_DONE on the owner window when finished.
     *
     * @param matrix        Receives the per-column-pair match counts.
     * @param bAutoMark     If true, also populates the key-missing arrays for diff highlighting.
     * @param bIn2file      If true, also checks for rows present in table 2 but not in table 1.
     * @param pbGreenClms1  Receives per-column "fully matched" flags for table 1.
     * @param pbGreenClms2  Receives per-column "fully matched" flags for table 2.
     * @param nEffMax       Receives the number of successfully key-matched row pairs.
     * @param bDoAutoMark   Receives the effective auto-mark flag (mirrors @p bAutoMark).
     */
    void firstPass(ComparisonMatrix& matrix,
                   bool bAutoMark, bool bIn2file,
                   std::vector<bool>& pbGreenClms1,
                   std::vector<bool>& pbGreenClms2,
                   int& nEffMax, bool& bDoAutoMark)
    {
        if (!m_bPrereq1valid) makePrereq1();
        if (!m_bPrereq2valid) makePrereq2();
        bDoAutoMark = bAutoMark;
        int prgHlpr = 0, prgHlpr0 = 0;
        char firstChar1, firstChar2;
        nEffMax = 0;
        matrix.clear(m_Table2.NumberOfColumns + 1, m_Table1.NumberOfColumns + 1);
        CString concatenatedKey1, concatenatedKey2;
        long keyRow2;
        int fchar1_y, fchar2_y;

        if (bAutoMark)
        {
            for (long i1 = m_Table1.FirstRowWithData; i1 <= m_Table1.NumberOfRows; i1++)
            {
                prgHlpr0 = 99 * i1 / m_Table1.NumberOfRows;
                if (prgHlpr0 > prgHlpr)
                {
                    prgHlpr = prgHlpr0;
                    ::PostMessage(m_hWnd, CM_UPDATE_PROGRESS, 0, prgHlpr);
                }
                concatenatedKey1 = m_pszKeyArr11[i1];
                if (m_Map2.Lookup(concatenatedKey1, (long&)keyRow2))
                {
                    nEffMax++;
                    fchar1_y = (i1 - 1) * m_Table1.NumberOfColumns;
                    for (int i3 = 1; i3 <= m_Table1.NumberOfColumns; i3++)
                    {
                        firstChar1 = m_pchMainArr1[fchar1_y + i3];
                        fchar2_y = (keyRow2 - 1) * m_Table2.NumberOfColumns;
                        for (int i4 = 1; i4 <= m_Table2.NumberOfColumns; i4++)
                        {
                            firstChar2 = m_pchMainArr2[fchar2_y + i4];
                            if (firstChar1 == firstChar2)
                            {
                                if (firstChar1 == 0 || (m_pExcel1->getCellValue(i3, i1) == m_pExcel2->getCellValue(i4, keyRow2)))
                                {
                                    matrix.increment(i4, i3);
                                }
                                else
                                {
                                    if (m_Table1.Columns[i3] == m_Table2.Columns[i4])
                                    {
                                        m_pchMainArr1[fchar1_y + i3] = 1;
                                        m_pchMainArr2[fchar2_y + i4] = 1;
                                    }
                                }
                            }
                            else
                            {
                                if (m_Table1.Columns[i3] == m_Table2.Columns[i4])
                                {
                                    m_pchMainArr1[fchar1_y + i3] = 1;
                                    m_pchMainArr2[fchar2_y + i4] = 1;
                                }
                            }
                        }
                    }
                }
                else
                {
                    m_pbKeyMissing1[i1] = true;
                }
            }
            if (bIn2file)
            {
                long keyRow1;
                prgHlpr = 0; prgHlpr0 = 0;
                for (long i1_2 = m_Table2.FirstRowWithData; i1_2 <= m_Table2.NumberOfRows; i1_2++)
                {
                    prgHlpr0 = 100 * i1_2 / m_Table2.NumberOfRows;
                    if (prgHlpr0 > prgHlpr)
                    {
                        prgHlpr = prgHlpr0;
                        ::PostMessage(m_hWnd, CM_UPDATE_PROGRESS, 0, prgHlpr);
                    }
                    concatenatedKey2 = m_pszKeyArr21[i1_2];
                    if (!m_Map1.Lookup(concatenatedKey2, (long&)keyRow1))
                        m_pbKeyMissing2[i1_2] = true;
                }
            }
            ::PostMessage(m_hWnd, CM_UPDATE_PROGRESS, 0, 100);
        }
        else
        {
            for (long i1 = m_Table1.FirstRowWithData; i1 <= m_Table1.NumberOfRows; i1++)
            {
                prgHlpr0 = 100 * i1 / m_Table1.NumberOfRows;
                if (prgHlpr0 > prgHlpr)
                {
                    prgHlpr = prgHlpr0;
                    ::PostMessage(m_hWnd, CM_UPDATE_PROGRESS, 0, prgHlpr);
                }
                concatenatedKey1 = m_pszKeyArr11[i1];
                if (m_Map2.Lookup(concatenatedKey1, (long&)keyRow2))
                {
                    nEffMax++;
                    fchar1_y = (i1 - 1) * m_Table1.NumberOfColumns;
                    for (int i3 = 1; i3 <= m_Table1.NumberOfColumns; i3++)
                    {
                        firstChar1 = m_pchMainArr1[fchar1_y + i3];
                        fchar2_y = (keyRow2 - 1) * m_Table2.NumberOfColumns;
                        for (int i4 = 1; i4 <= m_Table2.NumberOfColumns; i4++)
                        {
                            firstChar2 = m_pchMainArr2[fchar2_y + i4];
                            if (firstChar1 == firstChar2)
                            {
                                if (firstChar1 == 0 || (m_pExcel1->getCellValue(i3, i1) == m_pExcel2->getCellValue(i4, keyRow2)))
                                    matrix.increment(i4, i3);
                            }
                        }
                    }
                }
            }
        }

        pbGreenClms1.assign(m_Table1.NumberOfColumns + 2, false);
        pbGreenClms2.assign(m_Table2.NumberOfColumns + 2, false);
        for (int i_c = 1; i_c <= m_Table2.NumberOfColumns; i_c++)
        {
            for (int i_r = 1; i_r <= m_Table1.NumberOfColumns; i_r++)
            {
                if (matrix.get(i_c, i_r) == nEffMax)
                {
                    pbGreenClms1[i_r] = true;
                    pbGreenClms2[i_c] = true;
                }
            }
        }
        ::PostMessage(m_hWnd, CM_FIRSTPASS_DONE, 0, 0);
    }
    /** @brief Returns the column index of the @p key-th key for @p table (1 or 2). */
    int getNthKey(int table, int key) const
    {
        return (table == 1) ? m_KeyPair[key].tab1 : m_KeyPair[key].tab2;
    }
    /** @brief Overwrites the @p n-th key pair with new column indices. */
    void setNthKey(int n, int col1, int col2) { m_KeyPair[n].tab1 = col1; m_KeyPair[n].tab2 = col2; }
    /** @brief Appends a new key pair (one column from each table) to the active key list. */
    void pushKey(int col1, int col2)
    {
        m_KeyPair[m_nKeyPairCounter].tab1 = col1;
        m_KeyPair[m_nKeyPairCounter].tab2 = col2;
        m_nKeyPairCounter++;
    }
    /** @brief Removes all key pairs, resetting the key counter to zero. */
    void deleteAllKeys()
    {
        for (int i = 0; i < m_nKeyPairCounter; i++) { m_KeyPair[i].tab1 = 0; m_KeyPair[i].tab2 = 0; }
        m_nKeyPairCounter = 0;
    }
    /**
     * @brief Removes all key pairs that reference @p column in @p table (1 or 2).
     * @return Number of key pairs removed.
     */
    int deleteKey(int table, int column)
    {
        int rslt = 0;
        for (int i = 0; i < m_nKeyPairCounter; i++)
        {
            if ((table == 1 && m_KeyPair[i].tab1 == column) ||
                (table == 2 && m_KeyPair[i].tab2 == column))
            {
                deleteKeyAt(i--);
                rslt++;
            }
        }
        return rslt;
    }
    /** @brief Removes the key pair at position @p n, shifting subsequent pairs down. */
    void deleteKeyAt(int n)
    {
        for (int i = n; i < m_nKeyPairCounter; i++)
        {
            m_KeyPair[i].tab1 = m_KeyPair[i + 1].tab1;
            m_KeyPair[i].tab2 = m_KeyPair[i + 1].tab2;
        }
        m_nKeyPairCounter--;
    }
    /** @brief Inserts a key pair at position @p n, shifting subsequent pairs up. */
    void insertKeyAt(int n, int col1, int col2)
    {
        for (int i = m_nKeyPairCounter; i > n; i--)
        {
            m_KeyPair[i].tab1 = m_KeyPair[i - 1].tab1;
            m_KeyPair[i].tab2 = m_KeyPair[i - 1].tab2;
        }
        m_nKeyPairCounter++;
    }
    /** @brief Returns @c true if @p column in @p table (1 or 2) is part of any active key pair. */
    bool isThisAKey(int table, int column) const
    {
        for (int i = 0; i < m_nKeyPairCounter; i++)
        {
            if (table == 1 && m_KeyPair[i].tab1 == column) return true;
            if (table == 2 && m_KeyPair[i].tab2 == column) return true;
        }
        return false;
    }
    /** @brief Returns @c true if at least one key pair has been defined. */
    bool areThereAnyKeys() const { return m_nKeyPairCounter > 0; }
    /** @brief Returns the number of active key pairs. */
    int  getKeyPairCounter() const { return m_nKeyPairCounter; }

    // --- Data accessors ---
    /** @brief Returns the concatenated key string for row @p row in table 1. */
    CString getKeyStr1(int row) const { return m_pszKeyArr11[row]; }
    /** @brief Returns the concatenated key string for row @p row in table 2. */
    CString getKeyStr2(int row) const { return m_pszKeyArr21[row]; }
    /** @brief Returns @c true if row @p row in table 1 has no matching key in table 2. */
    bool    isKeyMissing1(int row) const { return m_pbKeyMissing1[row]; }
    /** @brief Returns @c true if row @p row in table 2 has no matching key in table 1. */
    bool    isKeyMissing2(int row) const { return m_pbKeyMissing2[row]; }
    /** @brief Returns @c true if column @p col in table 1 contains no data. */
    bool    isEmptyCol1(int col)   const { return m_pbEmptyClms1[col]; }
    /** @brief Returns @c true if column @p col in table 2 contains no data. */
    bool    isEmptyCol2(int col)   const { return m_pbEmptyClms2[col]; }
    /** @brief Returns the cached first character of cell (row, col) in table 1. */
    char    getMainChar1(int row, int col) const { return m_pchMainArr1[(row - 1) * m_Table1.NumberOfColumns + col]; }
    /** @brief Returns the cached first character of cell (row, col) in table 2. */
    char    getMainChar2(int row, int col) const { return m_pchMainArr2[(row - 1) * m_Table2.NumberOfColumns + col]; }
    /** @brief Returns a reference to the key-to-row lookup map for table 1. */
    CMap<CString, LPCTSTR, long, long>& getMap1() { return m_Map1; }
    /** @brief Returns a reference to the key-to-row lookup map for table 2. */
    CMap<CString, LPCTSTR, long, long>& getMap2() { return m_Map2; }

    NotUniqueKeys m_NotUniqueKeys1;
    NotUniqueKeys m_NotUniqueKeys2;
    bool          m_bUseIndexes = false;

private:
    // --- Internal helpers ---
    /** @brief Shared implementation for createKeyArrays1() and createKeyArrays2(). */
    int createKeyArraysImpl(int table)
    {
        const bool isT1 = (table == 1);
        NotUniqueKeys&                      notUniq    = isT1 ? m_NotUniqueKeys1 : m_NotUniqueKeys2;
        const Table&                        tbl        = isT1 ? m_Table1         : m_Table2;
        CMap<CString, LPCTSTR, long, long>& map        = isT1 ? m_Map1           : m_Map2;
        std::vector<CString>&               keyArr     = isT1 ? m_pszKeyArr11    : m_pszKeyArr21;
        std::vector<bool>&                  keyMissing = isT1 ? m_pbKeyMissing1  : m_pbKeyMissing2;
        ExcelConnector* const               pExcel     = isT1 ? m_pExcel1        : m_pExcel2;
        const UINT msgProgress = isT1 ? CM_UPDATE_PROGRESS : CM_UPDATE_PROGRESS2;
        const UINT msgDone     = isT1 ? CM_KEYS1_DONE      : CM_KEYS2_DONE;
        const int  dupCode     = isT1 ? 1                  : 2;

        notUniq = { 0, 0, L"" };
        long mapIdx;
        CString szdata;
        long idx = 0;
        map.RemoveAll();
        CString testdata;
        keyArr.assign(tbl.NumberOfRows + 2, L"");
        keyMissing.assign(tbl.NumberOfRows + 2, false);
        int prgHlpr = 0, prgHlpr0 = 0;
        for (int i_i = tbl.FirstRowWithData; i_i <= tbl.NumberOfRows; i_i++)
        {
            prgHlpr0 = 100 * i_i / tbl.NumberOfRows;
            if (prgHlpr0 > prgHlpr)
            {
                prgHlpr = prgHlpr0;
                ::PostMessage(m_hWnd, msgProgress, 0, prgHlpr);
            }
            szdata = L"";
            for (int k_i = 0; k_i < m_nKeyPairCounter; k_i++)
            {
                int nthKey = getNthKey(table, k_i);
                if (nthKey)
                    szdata += pExcel->getCellValue(nthKey, i_i);
            }
            keyMissing[i_i] = false;
            if (m_bUseIndexes)
            {
                idx = 0;
                do {
                    idx++;
                    testdata.Format(L"%s_idx%i", szdata, idx);
                } while (map.Lookup(testdata, (long&)mapIdx));
                szdata = testdata;
            }
            else
            {
                if (map.Lookup(szdata, (long&)mapIdx))
                {
                    notUniq = { i_i, mapIdx, szdata };
                    map.RemoveAll();
                    return dupCode;
                }
            }
            keyArr[i_i] = szdata;
            map.SetAt(szdata, i_i);
        }
        ::PostMessage(m_hWnd, msgDone, 0, 0);
        return 0;
    }

    /** @brief Shared implementation for checkKeysUniqueness1() and checkKeysUniqueness2(). */
    bool checkKeysUniquenessImpl(int table)
    {
        const bool isT1 = (table == 1);
        const Table&                tbl    = isT1 ? m_Table1      : m_Table2;
        const std::vector<CString>& keyArr = isT1 ? m_pszKeyArr11 : m_pszKeyArr21;
        const UINT msgProgress = isT1 ? CM_UPDATE_PROGRESS : CM_UPDATE_PROGRESS2;

        int prgHlpr = 0, prgHlpr0 = 0;
        CString szTaken_A, szTaken_B;
        for (int i0 = tbl.FirstRowWithData; i0 <= tbl.NumberOfRows; i0++)
        {
            prgHlpr0 = 100 * i0 / tbl.NumberOfRows;
            if (prgHlpr0 > prgHlpr)
            {
                prgHlpr = prgHlpr0;
                ::PostMessage(m_hWnd, msgProgress, 0, prgHlpr);
            }
            szTaken_A = keyArr[i0];
            for (int i1 = i0 + 1; i1 <= tbl.NumberOfRows; i1++)
            {
                szTaken_B = keyArr[i1];
                if (szTaken_A == szTaken_B)
                {
                    ::PostMessage(m_hWnd, msgProgress, 0, 100);
                    return false;
                }
            }
        }
        return true;
    }

    /** @brief Reads each cell in table 1 and stores its first character in m_pchMainArr1. */
    void makeCharArr1()
    {
        if (int arSize1 = (m_Table1.NumberOfColumns + 1) * (m_Table1.NumberOfRows + 1))
        {
            long prgHlpr0 = 0, prgHlpr = 0;
            m_pchMainArr1.assign(arSize1, 0);
            for (int i_c = 1; i_c <= m_Table1.NumberOfColumns; i_c++)
            {
                prgHlpr0 = 100 * i_c / m_Table1.NumberOfColumns;
                if (prgHlpr0 > prgHlpr) { prgHlpr = prgHlpr0; ::PostMessage(m_hWnd, CM_UPDATE_PROGRESS, 0, prgHlpr); }
                for (int i_r = 1; i_r <= m_Table1.NumberOfRows; i_r++)
                {
                    CString szdata = m_pExcel1->getCellValue(i_c, i_r);
                    m_pchMainArr1[(i_r - 1) * m_Table1.NumberOfColumns + i_c] = szdata.IsEmpty() ? 0 : szdata[0];
                }
            }
        }
        ::PostMessage(m_hWnd, CM_UPDATE_PROGRESS, 0, 100);
    }

    /** @brief Reads each cell in table 2 and stores its first character in m_pchMainArr2. */
    void makeCharArr2()
    {
        if (int arSize2 = (m_Table2.NumberOfColumns + 1) * (m_Table2.NumberOfRows + 1))
        {
            long prgHlpr0 = 0, prgHlpr = 0;
            m_pchMainArr2.assign(arSize2, 0);
            for (int i_c = 1; i_c <= m_Table2.NumberOfColumns; i_c++)
            {
                prgHlpr0 = 100 * i_c / m_Table2.NumberOfColumns;
                if (prgHlpr0 > prgHlpr) { prgHlpr = prgHlpr0; ::PostMessage(m_hWnd, CM_UPDATE_PROGRESS2, 0, prgHlpr); }
                for (int i_r = 1; i_r <= m_Table2.NumberOfRows; i_r++)
                {
                    CString szdata = m_pExcel2->getCellValue(i_c, i_r);
                    m_pchMainArr2[(i_r - 1) * m_Table2.NumberOfColumns + i_c] = szdata.IsEmpty() ? 0 : szdata[0];
                }
            }
        }
        ::PostMessage(m_hWnd, CM_UPDATE_PROGRESS2, 0, 100);
    }

    /** @brief Determines which columns of table 1 are entirely empty and flags them in m_pbEmptyClms1. */
    void checkEmptiness1()
    {
        m_pbEmptyClms1.assign(m_Table1.NumberOfColumns + 2, true);
        long prgHlpr0 = 0, prgHlpr = 0;
        for (int i_c = 1; i_c <= m_Table1.NumberOfColumns; i_c++)
        {
            prgHlpr0 = 100 * i_c / m_Table1.NumberOfColumns;
            if (prgHlpr0 > prgHlpr + 10) { prgHlpr = prgHlpr0; ::PostMessage(m_hWnd, CM_UPDATE_PROGRESS, 0, prgHlpr); }
            for (int i_r = m_Table1.FirstRowWithData; i_r <= m_Table1.NumberOfRows; i_r++)
            {
                if (m_pchMainArr1[(i_r - 1) * m_Table1.NumberOfColumns + i_c]) { m_pbEmptyClms1[i_c] = false; break; }
            }
        }
        ::PostMessage(m_hWnd, CM_UPDATE_PROGRESS, 0, 100);
    }

    /** @brief Determines which columns of table 2 are entirely empty and flags them in m_pbEmptyClms2. */
    void checkEmptiness2()
    {
        m_pbEmptyClms2.assign(m_Table2.NumberOfColumns + 2, true);
        long prgHlpr0 = 0, prgHlpr = 0;
        for (int i_c = 1; i_c <= m_Table2.NumberOfColumns; i_c++)
        {
            prgHlpr0 = 100 * i_c / m_Table2.NumberOfColumns;
            if (prgHlpr0 > prgHlpr + 10) { prgHlpr = prgHlpr0; ::PostMessage(m_hWnd, CM_UPDATE_PROGRESS2, 0, prgHlpr); }
            for (int i_r = m_Table2.FirstRowWithData; i_r <= m_Table2.NumberOfRows; i_r++)
            {
                if (m_pchMainArr2[(i_r - 1) * m_Table2.NumberOfColumns + i_c]) { m_pbEmptyClms2[i_c] = false; break; }
            }
        }
        ::PostMessage(m_hWnd, CM_UPDATE_PROGRESS2, 0, 100);
    }

    // --- Internal state ---
    HWND             m_hWnd    = nullptr;
    ExcelConnector*  m_pExcel1 = nullptr;
    ExcelConnector*  m_pExcel2 = nullptr;
    Table            m_Table1  = {};
    Table            m_Table2  = {};

    bool m_bPrereq1valid = false;
    bool m_bPrereq2valid = false;

    std::vector<char>    m_pchMainArr1;
    std::vector<char>    m_pchMainArr2;
    std::vector<bool>    m_pbEmptyClms1;
    std::vector<bool>    m_pbEmptyClms2;
    std::vector<CString> m_pszKeyArr11;
    std::vector<CString> m_pszKeyArr21;
    std::vector<bool>    m_pbKeyMissing1;
    std::vector<bool>    m_pbKeyMissing2;

    CMap<CString, LPCTSTR, long, long> m_Map1;
    CMap<CString, LPCTSTR, long, long> m_Map2;

    KeyPair m_KeyPair[256] = {};
    int     m_nKeyPairCounter = 0;
};
