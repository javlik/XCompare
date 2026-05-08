#pragma once
#include "Constants.h"
#include "TableData.h"
#include "ExcelConnector.h"
#include <vector>
#include <map>
#include <cmath>

/**
 * @brief Encapsulates the automatic key-suggestion algorithm.
 *
 * Analyses each column of both tables for uniqueness (entropy) and searches
 * for the best combination of columns that could serve as a join key.
 * Results are stored in @c m_PossibleKeys1/@c m_PossibleKeys2 and exposed
 * via accessor methods so CChildView can push the winner into @c ComparisonEngine.
 *
 * Typical usage:
 * @code
 *   KeyFinder m_keyFinder;
 *   m_keyFinder.init(hWnd, m_excel1, m_excel2, m_Table1, m_Table2);
 *   m_keyFinder.setComplexity(100000);
 *   // --- in worker thread ---
 *   m_keyFinder.suggestKeys1();
 *   // --- after thread ---
 *   m_keyFinder.mutualCheck();
 *   if (m_keyFinder.usePossibleKeys(...))
 *       // push into ComparisonEngine
 * @endcode
 */
class KeyFinder
{
public:
    // --- Initialisation ---
    /** @brief Initialises the finder with window handle, Excel connections and table descriptors. */
    void init(HWND hWnd,
              ExcelConnector& excel1, ExcelConnector& excel2,
              const Table& table1, const Table& table2)
    {
        m_hWnd    = hWnd;
        m_pExcel1 = &excel1;
        m_pExcel2 = &excel2;
        m_Table1  = table1;
        m_Table2  = table2;
        m_nCheckedKeys1.assign(MAX_ATTEMPTS + 2, 0ULL);
        m_nCheckedKeys2.assign(MAX_ATTEMPTS + 2, 0ULL);
    }

    /** @brief Refreshes the table descriptors after a sheet selection change. */
    void setTables(const Table& table1, const Table& table2)
    {
        m_Table1 = table1;
        m_Table2 = table2;
    }

    /** @brief Sets the maximum number of candidate combinations to examine per table. */
    void setComplexity(int complexity) { m_nComplexity = complexity; }

    // --- Key-suggestion main entry points (run in worker threads) ---
    /**
     * @brief Searches table 1 for candidate key column combinations.
     *
     * Computes per-column entropy, then iterates through combinations of up to
     * @c SUGKEYS columns, checking each for uniqueness. Stores successful candidates
     * in @c m_PossibleKeys1. Posts @c CM_GATHERING1_DONE when finished.
     */
    void suggestKeys1() { suggestKeysImpl(1); }

    /**
     * @brief Searches table 2 for candidate key column combinations.
     *
     * Mirror of @c suggestKeys1() for the second table.
     * Posts @c CM_GATHERING2_DONE when finished.
     */
    void suggestKeys2() { suggestKeysImpl(2); }

    // --- Support methods called from CChildView ---
    /** @brief Resets all candidate key arrays and counters for both tables. */
    void clearPossibleKeys()
    {
        for (int i = 0; i < 255; i++)
            for (int ii = 0; ii < SUGKEYS; ii++)
            {
                m_PossibleKeys1[i].k[ii] = 0;
                m_PossibleKeys2[i].k[ii] = 0;
            }
        m_nPossibleKeyCounter1 = 0;
        m_nPossibleKeyCounter2 = 0;
    }

    /**
     * @brief Cross-checks one table-1 candidate key set against all matching table-2 sets.
     * @param tab1 Index into @c m_PossibleKeys1 to test.
     * @return The best match percentage (0–100) found, or -1 if the candidate is empty.
     */
    int checkKeys(int tab1)
    {
        CString szdata;
        int prgHlpr = 0, prgHlpr0 = 0;
        m_mapTmpMap1.clear();
        int order1 = getNumberOfPossibleKeys(1, SUGKEYS, tab1);
        if (order1 > 0)
        {
            for (int i_i = m_Table1.FirstRowWithData; i_i <= m_Table1.NumberOfRows; i_i++)
            {
                prgHlpr0 = 100 * i_i / m_Table1.NumberOfRows;
                if (prgHlpr0 > prgHlpr)
                {
                    prgHlpr = prgHlpr0;
                    ::PostMessage(m_hWnd, CM_UPDATE_PROGRESS2,    0, prgHlpr);
                    ::PostMessage(m_hWnd, CM_UPDATE_KEYPROGRESS2, 0, prgHlpr);
                }
                szdata = L"";
                for (int i_j = 0; i_j <= order1; i_j++)
                {
                    if (m_PossibleKeys1[tab1].k[i_j])
                        szdata += m_pExcel1->getCellValue(m_PossibleKeys1[tab1].k[i_j], i_i);
                }
                m_mapTmpMap1[szdata] = i_i;
            }
        }
        else
        {
            return -1;
        }
        int tab2 = 0;
        while (getNumberOfPossibleKeys(2, SUGKEYS, tab2) < order1)
            tab2++;
        long rslt = 0;
        while (getNumberOfPossibleKeys(2, SUGKEYS, tab2) == order1)
        {
            long found = 0;
            for (int i_i = m_Table2.FirstRowWithData; i_i <= m_Table2.NumberOfRows; i_i++)
            {
                prgHlpr0 = 100 * i_i / m_Table2.NumberOfRows;
                if (prgHlpr0 > prgHlpr)
                {
                    prgHlpr = prgHlpr0;
                    ::PostMessage(m_hWnd, CM_UPDATE_PROGRESS2,    0, prgHlpr);
                    ::PostMessage(m_hWnd, CM_UPDATE_KEYPROGRESS2, 0, prgHlpr);
                }
                szdata = L"";
                for (int i_k = 0; i_k <= order1; i_k++)
                {
                    if (m_PossibleKeys2[tab2].k[i_k])
                        szdata += m_pExcel2->getCellValue(m_PossibleKeys2[tab2].k[i_k], i_i);
                }
                if (m_mapTmpMap1.count(szdata))
                    found++;
            }
            rslt = 100 * found / min(m_Table2.NumberOfRows - m_Table2.FirstRowWithData,
                                     m_Table1.NumberOfRows - m_Table1.FirstRowWithData);
            if (rslt > m_BestKeyComb.rating)
            {
                m_BestKeyComb.pk1    = tab1;
                m_BestKeyComb.pk2    = tab2;
                m_BestKeyComb.rating = (int)rslt;
                m_BestKeyComb.cnt    = found;
            }
            tab2++;
        }
        ::PostMessage(m_hWnd, CM_UPDATE_KEYPROGRESS1, 0, 0);
        ::PostMessage(m_hWnd, CM_UPDATE_KEYPROGRESS2, 0, 0);
        return m_BestKeyComb.rating;
    }

    // --- Accessors used by CChildView ---
    /** @brief Returns the @p idx-th key column index for candidate @p item in table 1. */
    int  getPossibleKey1(int item, int idx) const { return m_PossibleKeys1[item].k[idx]; }
    /** @brief Returns the @p idx-th key column index for candidate @p item in table 2. */
    int  getPossibleKey2(int item, int idx) const { return m_PossibleKeys2[item].k[idx]; }

    /**
     * @brief Returns the number of key columns in the best candidate pair found so far.
     *
     * Counts the leading non-zero entries in the best pair's @c PossibleKeys array.
     */
    // Number of non-zero keys in the best candidate pair
    int getNumberOfPossibleKeys() const
    {
        for (int tmp_i = 1; tmp_i < 255; tmp_i++)
        {
            if (m_PossibleKeys1[m_BestKeyComb.pk1].k[tmp_i] == 0 &&
                m_PossibleKeys2[m_BestKeyComb.pk2].k[tmp_i] == 0)
                return tmp_i;
        }
        return 0;
    }

    /**
     * @brief Returns the number of non-zero key slots (up to @p order) in candidate @p item for @p table (1 or 2).
     */
    // Number of non-zero key slots (up to 'order') in candidate 'item' for the given table
    int getNumberOfPossibleKeys(int table, int order, int item) const
    {
        int cnt = 0;
        if (table == 1)
            for (int i = 0; i < order; i++)
                cnt += sgn(m_PossibleKeys1[item].k[i]);
        else
            for (int i = 0; i < order; i++)
                cnt += sgn(m_PossibleKeys2[item].k[i]);
        return cnt;
    }

    /** @brief Returns a copy of the best key combination found by @c checkKeys(). */
    BestKeyComb getBestKeyComb()         const { return m_BestKeyComb; }
    /** @brief Resets the best-key-combination record. */
    void        resetBestKeyComb()             { m_BestKeyComb = {}; }
    /** @brief Returns the number of candidate key sets found for table 1. */
    int         getPossibleKeyCounter1() const { return m_nPossibleKeyCounter1; }
    /** @brief Returns the number of candidate key sets found for table 2. */
    int         getPossibleKeyCounter2() const { return m_nPossibleKeyCounter2; }

private:
    // --- Internal helpers ---

    /**
     * @brief Shared implementation for suggestKeys1() and suggestKeys2().
     *
     * All per-table differences (member arrays, message IDs, resource strings) are
     * resolved via local references/pointers at the top, so the algorithm body is
     * written only once.
     * @param table 1 = operate on table 1, 2 = operate on table 2.
     */
    void suggestKeysImpl(int table)
    {
        // --- Message IDs that differ between the two tables ---

        // --- Per-table references (ternary on lvalues of the same type yields a reference) ---
        const bool isT1 = (table == 1);
        const UINT msgProgress = isT1 ? CM_UPDATE_PROGRESS      : CM_UPDATE_PROGRESS2;
        const UINT msgKeyProg  = isT1 ? CM_UPDATE_KEYPROGRESS1   : CM_UPDATE_KEYPROGRESS2;
        const UINT msgDone     = isT1 ? CM_GATHERING1_DONE       : CM_GATHERING2_DONE;
        const UINT idNoSheet   = isT1 ? IDS_NO_SHEET_SELCTD_IN_FRST : IDS_NO_SHEET_SELCTD_IN_SCND;

        const Table&                     tbl      = isT1 ? m_Table1               : m_Table2;
        ExcelConnector* const            pExcel   = isT1 ? m_pExcel1              : m_pExcel2;
        long* const                      invEnt   = isT1 ? m_nInvEntropy1         : m_nInvEntropy2;
        int*  const                      exKeys   = isT1 ? m_nExaminedKeys1       : m_nExaminedKeys2;
        PossibleKeys*                    posKeys  = isT1 ? m_PossibleKeys1        : m_PossibleKeys2;
        int&                             posKeyCnt= isT1 ? m_nPossibleKeyCounter1 : m_nPossibleKeyCounter2;
        std::map<CString, long>&         tmpMap   = isT1 ? m_mapTmpMap1           : m_mapTmpMap2;
        int&                             chKeyCnt = isT1 ? m_nCheckedKeysCounter1 : m_nCheckedKeysCounter2;
        std::vector<unsigned long long>& chKeys   = isT1 ? m_nCheckedKeys1        : m_nCheckedKeys2;

        // --- Initialise state ---
        int attempts = 0;
        bool alreadyChecked = false;
        chKeyCnt = 0;
        for (int i = 0; i < SUGKEYS; i++)
            chKeys[i] = 0;
        int prgHlpr = 0, prgHlpr0 = 0;
        posKeyCnt = 0;
        for (int i = 0; i < 255; i++)
            invEnt[i] = 0;
        for (int i = 0; i <= SUGKEYS + 1; i++)
            exKeys[i] = 0;

        // --- Compute per-column inverse entropy ---
        CString szdata;
        for (int i_h = 1; i_h <= tbl.NumberOfColumns; i_h++)
        {
            tmpMap.clear();
            for (int i_i = tbl.FirstRowWithData; i_i <= tbl.NumberOfRows; i_i++)
            {
                szdata = pExcel->getCellValue(i_h, i_i);
                if (tmpMap.find(szdata) == tmpMap.end())
                {
                    tmpMap[szdata] = i_i;
                    invEnt[i_h]++;
                }
            }
        }
        CalculateEntropyRank(table);

        // --- Main search loop ---
        if (tbl.NumberOfRows > 0)
        {
            int foundKeysSet = 10;
            while (true)
            {
                prgHlpr0 = attempts % 97;
                if (prgHlpr0 > prgHlpr)
                {
                    prgHlpr = prgHlpr0;
                    ::PostMessage(m_hWnd, msgProgress, 0, prgHlpr);
                    ::PostMessage(m_hWnd, msgKeyProg,  0, prgHlpr);
                }
                if (is2BExaminedOnce(table, SUGKEYS - 1))
                {
                    alreadyChecked = getSimilarKeyProbability(table, SUGKEYS);
                    if (!alreadyChecked)
                        foundKeysSet = createTempKeyArrays(table);
                }
                else
                {
                    foundKeysSet = 4; // Low entropy of key indexes
                }
                if (foundKeysSet == 0)
                {
                    for (int tmp_i = 0; tmp_i < SUGKEYS; tmp_i++)
                        posKeys[posKeyCnt].k[tmp_i] = getNthEntropy(table, exKeys[tmp_i]);
                    sortExaminedKeys(table);
                    posKeyCnt++;
                }
                if (attempts++ > m_nComplexity || posKeyCnt > tbl.NumberOfColumns)
                    break;
                int e_i = SUGKEYS - 1;
                while (e_i >= 0)
                {
                    if (exKeys[e_i] >= tbl.NumberOfColumns)
                    {
                        exKeys[e_i] = 0;
                        --e_i;
                    }
                    else
                    {
                        ++exKeys[e_i];
                        break;
                    }
                }
            }
        }
        else
        {
            ::MessageBox(m_hWnd, CMsg(idNoSheet), nullptr, MB_OK);
        }
        ::PostMessage(m_hWnd, msgDone, 0, 0);
    }

    /**
     * @brief Builds a temporary uniqueness map for the current examined-key combination.
     * @param table 1 or 2.
     * @return 0 = combination is unique, 1 = all examined indices are zero, 2 = duplicate found.
     */
    int createTempKeyArrays(int table)
    {
        const bool isT1 = (table == 1);
        int*  const              exKeys = isT1 ? m_nExaminedKeys1 : m_nExaminedKeys2;
        const Table&             tbl    = isT1 ? m_Table1         : m_Table2;
        ExcelConnector* const    pExcel = isT1 ? m_pExcel1        : m_pExcel2;
        std::map<CString, long>& tmpMap = isT1 ? m_mapTmpMap1     : m_mapTmpMap2;

        CString szdata;
        tmpMap.clear();
        if (sumExaminedKeys(table, SUGKEYS - 1) > 0)
        {
            for (int i_i = tbl.FirstRowWithData; i_i <= tbl.NumberOfRows; i_i++)
            {
                szdata = L"";
                for (int k_i = 0; k_i < SUGKEYS; k_i++)
                {
                    if (exKeys[k_i])
                        szdata += pExcel->getCellValue(getNthEntropy(table, exKeys[k_i]), i_i);
                }
                if (tmpMap.find(szdata) != tmpMap.end())
                {
                    tmpMap.clear();
                    return 2;
                }
                tmpMap[szdata] = i_i;
            }
        }
        else
        {
            return 1;
        }
        return 0;
    }

    /** @brief Compacts a just-recorded PossibleKeys entry by moving non-zero indices to the front. */
    void sortExaminedKeys(int table)
    {
        int nonzerosNr = 0;
        if (table == 1)
        {
            for (int i_i = 0; i_i < SUGKEYS; i_i++)
            {
                if (m_PossibleKeys1[m_nPossibleKeyCounter1].k[i_i] > 0 && nonzerosNr < i_i)
                {
                    m_PossibleKeys1[m_nPossibleKeyCounter1].k[nonzerosNr] = m_PossibleKeys1[m_nPossibleKeyCounter1].k[i_i];
                    m_PossibleKeys1[m_nPossibleKeyCounter1].k[i_i] = 0;
                    nonzerosNr++;
                }
            }
        }
        else
        {
            for (int i_i = 0; i_i < SUGKEYS; i_i++)
            {
                if (m_PossibleKeys2[m_nPossibleKeyCounter2].k[i_i] > 0 && nonzerosNr < i_i)
                {
                    m_PossibleKeys2[m_nPossibleKeyCounter2].k[nonzerosNr] = m_PossibleKeys2[m_nPossibleKeyCounter2].k[i_i];
                    m_PossibleKeys2[m_nPossibleKeyCounter2].k[i_i] = 0;
                    nonzerosNr++;
                }
            }
        }
    }

    /** @brief Returns the sum of the first @p nmax+1 examined key indices for @p table (1 or 2). */
    int sumExaminedKeys(int table, int nmax) const
    {
        int rslt = 0;
        if (table == 1)
            for (int tmp_i = 0; tmp_i <= nmax; tmp_i++)
                rslt += m_nExaminedKeys1[tmp_i];
        else
            for (int tmp_i = 0; tmp_i <= nmax; tmp_i++)
                rslt += m_nExaminedKeys2[tmp_i];
        return rslt;
    }

    /** @brief Returns @c true if the first @p max+1 examined key indices contain no duplicates. */
    bool is2BExaminedOnce(int table, int max) const
    {
        if (table == 1)
        {
            for (int tmp_i0 = 0; tmp_i0 <= max; tmp_i0++)
            {
                if (m_nExaminedKeys1[tmp_i0] > 0)
                    for (int tmp_i1 = tmp_i0 + 1; tmp_i1 <= max; tmp_i1++)
                        if (m_nExaminedKeys1[tmp_i0] == m_nExaminedKeys1[tmp_i1])
                            return false;
            }
        }
        else
        {
            for (int tmp_i0 = 0; tmp_i0 <= max; tmp_i0++)
            {
                if (m_nExaminedKeys2[tmp_i0] > 0)
                    for (int tmp_i1 = tmp_i0 + 1; tmp_i1 <= max; tmp_i1++)
                        if (m_nExaminedKeys2[tmp_i0] == m_nExaminedKeys2[tmp_i1])
                            return false;
            }
        }
        return true;
    }

    /** @brief Returns @c true if the current examined-keys combination has already been tested (deduplication). */
    bool getSimilarKeyProbability(int table, int max)
    {
        unsigned long long ullTest = 0;
        if (table == 1)
        {
            for (int tmp_i = 0; tmp_i < SUGKEYS; tmp_i++)
                if (m_nExaminedKeys1[tmp_i])
                    ullTest += (unsigned long long)pow(2.0, m_nExaminedKeys1[tmp_i]);
            for (int tmp_i0 = 0; tmp_i0 <= m_nCheckedKeysCounter1; tmp_i0++)
                if (m_nCheckedKeys1[tmp_i0] == ullTest)
                    return true;
            if (m_nCheckedKeysCounter1 < m_nComplexity)
            {
                m_nCheckedKeys1[m_nCheckedKeysCounter1] = ullTest;
                m_nCheckedKeysCounter1++;
            }
            else
            {
                return true;
            }
        }
        else
        {
            for (int tmp_i = 0; tmp_i < SUGKEYS; tmp_i++)
                if (m_nExaminedKeys2[tmp_i])
                    ullTest += (unsigned long long)pow(2.0, m_nExaminedKeys2[tmp_i]);
            for (int tmp_i0 = 0; tmp_i0 <= m_nCheckedKeysCounter2; tmp_i0++)
                if (m_nCheckedKeys2[tmp_i0] == ullTest)
                    return true;
            if (m_nCheckedKeysCounter2 < m_nComplexity)
            {
                m_nCheckedKeys2[m_nCheckedKeysCounter2] = ullTest;
                m_nCheckedKeysCounter2++;
            }
            else
            {
                return true;
            }
        }
        return false;
    }

    /** @brief Returns the column index ranked @p n-th by entropy for @p table (1 or 2). */
    int getNthEntropy(int table, int n) const
    {
        return (table == 1) ? m_nSortedEntropy1[n] : m_nSortedEntropy2[n];
    }

    /** @brief Sorts all columns by descending uniqueness and stores the ranking in @c m_nSortedEntropy1/2. */
    int CalculateEntropyRank(int table)
    {
        int stored = 0;
        if (table == 1)
        {
            for (int i0 = 0; i0 < 255; i0++)
                m_nSortedEntropy1[i0] = 0;
            while (stored < m_Table1.NumberOfColumns)
            {
                int hlpr_index = 0, hlpr_value = 0;
                for (int i1 = 1; i1 <= m_Table1.NumberOfColumns; i1++)
                {
                    if (m_nInvEntropy1[i1] >= hlpr_value && !isEntropyStored(1, i1, stored))
                    {
                        hlpr_value = (int)m_nInvEntropy1[i1];
                        hlpr_index = i1;
                    }
                }
                if (hlpr_index > 0)
                {
                    stored++;
                    m_nSortedEntropy1[stored] = hlpr_index;
                }
            }
        }
        else
        {
            for (int i0 = 0; i0 < 255; i0++)
                m_nSortedEntropy2[i0] = 0;
            while (stored < m_Table2.NumberOfColumns)
            {
                int hlpr_index = 0, hlpr_value = 0;
                for (int i2 = 1; i2 <= m_Table2.NumberOfColumns; i2++)
                {
                    if (m_nInvEntropy2[i2] >= hlpr_value && !isEntropyStored(2, i2, stored))
                    {
                        hlpr_value = (int)m_nInvEntropy2[i2];
                        hlpr_index = i2;
                    }
                }
                if (hlpr_index > 0)
                {
                    stored++;
                    m_nSortedEntropy2[stored] = hlpr_index;
                }
            }
        }
        return 0;
    }

    /** @brief Returns @c true if column @p clm already appears in the top @p max sorted-entropy entries for @p table. */
    bool isEntropyStored(int table, int clm, int max) const
    {
        if (table == 1)
            for (int i = 1; i <= max; i++)
                if (m_nSortedEntropy1[i] == clm) return true;
        else
            for (int i = 1; i <= max; i++)
                if (m_nSortedEntropy2[i] == clm) return true;
        return false;
    }

    // --- External references ---
    HWND            m_hWnd    = nullptr;
    ExcelConnector* m_pExcel1 = nullptr;
    ExcelConnector* m_pExcel2 = nullptr;
    Table           m_Table1  = {};
    Table           m_Table2  = {};
    int             m_nComplexity = 100000;

    // --- Key-suggestion state ---
    long m_nInvEntropy1[256]    = {};
    long m_nInvEntropy2[256]    = {};
    int  m_nSortedEntropy1[256] = {};
    int  m_nSortedEntropy2[256] = {};

    int m_nExaminedKeys1[SUGKEYS + 4] = {};
    int m_nExaminedKeys2[SUGKEYS + 4] = {};

    std::vector<unsigned long long> m_nCheckedKeys1;
    std::vector<unsigned long long> m_nCheckedKeys2;
    int m_nCheckedKeysCounter1 = 0;
    int m_nCheckedKeysCounter2 = 0;

    PossibleKeys m_PossibleKeys1[256] = {};
    PossibleKeys m_PossibleKeys2[256] = {};
    int          m_nPossibleKeyCounter1 = 0;
    int          m_nPossibleKeyCounter2 = 0;
    BestKeyComb  m_BestKeyComb = {};

    std::map<CString, long> m_mapTmpMap1;
    std::map<CString, long> m_mapTmpMap2;
};
