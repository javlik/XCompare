#pragma once
#include "TableData.h"
#include "ComparisonMatrix.h"
#include "Constants.h"
#include "ExcelConnector.h"
#include <vector>
#include <map>

// ComparisonEngine encapsulates the comparison algorithm data and methods
// that were previously embedded directly in CChildView.
//
// It holds all per-comparison data arrays, key structures and entropy tables.
// It needs a HWND for progress PostMessage calls and references to the two
// ExcelConnector objects so it can read cell values.
//
// Usage:
//   ComparisonEngine m_engine;
//   m_engine.init(hWnd, m_excel1, m_excel2, m_Table1, m_Table2);
//   m_engine.makePrereq1();
//   m_engine.firstPass();

class ComparisonEngine
{
public:
    // --- Initialisation ---
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

    // Refresh the table descriptors (called after sheet selection)
    void setTables(const Table& table1, const Table& table2)
    {
        m_Table1 = table1;
        m_Table2 = table2;
    }

    // --- Prerequisite building ---
    void makePrereq1()
    {
        m_bPrereq1valid = false;
        m_pchMainArr1.assign((m_Table1.NumberOfColumns + 1) * (m_Table1.NumberOfRows + 1), 0);
        makeCharArr1();
        checkEmptiness1();
        m_bPrereq1valid = true;
    }

    void makePrereq2()
    {
        m_bPrereq2valid = false;
        m_pchMainArr2.assign((m_Table2.NumberOfColumns + 1) * (m_Table2.NumberOfRows + 1), 0);
        makeCharArr2();
        checkEmptiness2();
        m_bPrereq2valid = true;
    }

    bool isPrereq1Valid() const { return m_bPrereq1valid; }
    bool isPrereq2Valid() const { return m_bPrereq2valid; }
    void invalidatePrereq1()    { m_bPrereq1valid = false; }
    void invalidatePrereq2()    { m_bPrereq2valid = false; }

    // --- Key arrays ---
    // Returns 0 on success, 1 if keys in table 1 are not unique, 2 for table 2.
    int createKeyArrays1()
    {
        m_NotUniqueKeys1 = { 0, 0, L"" };
        long mapIdx;
        CString szdata;
        long idx = 0;
        m_Map1.RemoveAll();
        CString testdata;
        m_pszKeyArr11.assign(m_Table1.NumberOfRows + 2, L"");
        m_pbKeyMissing1.assign(m_Table1.NumberOfRows + 2, false);
        int prgHlpr = 0, prgHlpr0 = 0;
        for (int i_i = m_Table1.FirstRowWithData; i_i <= m_Table1.NumberOfRows; i_i++)
        {
            prgHlpr0 = 100 * i_i / m_Table1.NumberOfRows;
            if (prgHlpr0 > prgHlpr)
            {
                prgHlpr = prgHlpr0;
                ::PostMessage(m_hWnd, CM_UPDATE_PROGRESS, 0, prgHlpr);
            }
            szdata = L"";
            for (int k_i = 0; k_i < m_nKeyPairCounter; k_i++)
            {
                int nthKey = getNthKey(1, k_i);
                if (nthKey)
                    szdata += m_pExcel1->getCellValue(nthKey, i_i);
            }
            m_pbKeyMissing1[i_i] = false;
            if (m_bUseIndexes)
            {
                idx = 0;
                do {
                    idx++;
                    testdata.Format(L"%s_idx%i", szdata, idx);
                } while (m_Map1.Lookup(testdata, (long&)mapIdx));
                szdata = testdata;
            }
            else
            {
                if (m_Map1.Lookup(szdata, (long&)mapIdx))
                {
                    m_NotUniqueKeys1 = { i_i, mapIdx, szdata };
                    m_Map1.RemoveAll();
                    return 1;
                }
            }
            m_pszKeyArr11[i_i] = szdata;
            m_Map1.SetAt(szdata, i_i);
        }
        ::PostMessage(m_hWnd, CM_UPDATE_PROGRESS, 0, 1000);
        return 0;
    }

    int createKeyArrays2()
    {
        m_NotUniqueKeys2 = { 0, 0, L"" };
        long mapIdx;
        CString szdata;
        long idx = 0;
        m_Map2.RemoveAll();
        CString testdata;
        m_pszKeyArr21.assign(m_Table2.NumberOfRows + 2, L"");
        m_pbKeyMissing2.assign(m_Table2.NumberOfRows + 2, false);
        int prgHlpr = 0, prgHlpr0 = 0;
        for (int i_i = m_Table2.FirstRowWithData; i_i <= m_Table2.NumberOfRows; i_i++)
        {
            prgHlpr0 = 100 * i_i / m_Table2.NumberOfRows;
            if (prgHlpr0 > prgHlpr)
            {
                prgHlpr = prgHlpr0;
                ::PostMessage(m_hWnd, CM_UPDATE_PROGRESS2, 0, prgHlpr);
            }
            szdata = L"";
            for (int k_i = 0; k_i < m_nKeyPairCounter; k_i++)
            {
                int nthKey = getNthKey(2, k_i);
                if (nthKey)
                    szdata += m_pExcel2->getCellValue(nthKey, i_i);
            }
            m_pbKeyMissing2[i_i] = false;
            if (m_bUseIndexes)
            {
                idx = 0;
                do {
                    idx++;
                    testdata.Format(L"%s_idx%i", szdata, idx);
                } while (m_Map2.Lookup(testdata, (long&)mapIdx));
                szdata = testdata;
            }
            else
            {
                if (m_Map2.Lookup(szdata, (long&)mapIdx))
                {
                    m_NotUniqueKeys2 = { i_i, mapIdx, szdata };
                    m_Map2.RemoveAll();
                    return 2;
                }
            }
            m_pszKeyArr21[i_i] = szdata;
            m_Map2.SetAt(szdata, i_i);
        }
        ::PostMessage(m_hWnd, CM_UPDATE_PROGRESS2, 0, 1000);
        return 0;
    }

    bool checkKeysUniqueness1()
    {
        int prgHlpr = 0, prgHlpr0 = 0;
        CString szTaken_A, szTaken_B;
        for (int i0 = m_Table1.FirstRowWithData; i0 <= m_Table1.NumberOfRows; i0++)
        {
            prgHlpr0 = 100 * i0 / m_Table1.NumberOfRows;
            if (prgHlpr0 > prgHlpr)
            {
                prgHlpr = prgHlpr0;
                ::PostMessage(m_hWnd, CM_UPDATE_PROGRESS, 0, prgHlpr);
            }
            szTaken_A = m_pszKeyArr11[i0];
            for (int i1 = i0 + 1; i1 <= m_Table1.NumberOfRows; i1++)
            {
                szTaken_B = m_pszKeyArr11[i1];
                if (szTaken_A == szTaken_B)
                {
                    ::PostMessage(m_hWnd, CM_UPDATE_PROGRESS, 0, 100);
                    return false;
                }
            }
        }
        return true;
    }

    bool checkKeysUniqueness2()
    {
        int prgHlpr = 0, prgHlpr0 = 0;
        CString szTaken_A, szTaken_B;
        for (int i0 = m_Table2.FirstRowWithData; i0 <= m_Table2.NumberOfRows; i0++)
        {
            prgHlpr0 = 100 * i0 / m_Table2.NumberOfRows;
            if (prgHlpr0 > prgHlpr)
            {
                prgHlpr = prgHlpr0;
                ::PostMessage(m_hWnd, CM_UPDATE_PROGRESS2, 0, prgHlpr);
            }
            szTaken_A = m_pszKeyArr21[i0];
            for (int i1 = i0 + 1; i1 <= m_Table2.NumberOfRows; i1++)
            {
                szTaken_B = m_pszKeyArr21[i1];
                if (szTaken_A == szTaken_B)
                {
                    ::PostMessage(m_hWnd, CM_UPDATE_PROGRESS2, 0, 100);
                    return false;
                }
            }
        }
        return true;
    }

    // --- First pass (main comparison algorithm) ---
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
        ::PostMessage(m_hWnd, CM_UPDATE_PROGRESS, 0, 1000);
    }

    // --- Key management ---
    int getNthKey(int table, int key) const
    {
        return (table == 1) ? m_KeyPair[key].tab1 : m_KeyPair[key].tab2;
    }
    void setNthKey(int n, int col1, int col2) { m_KeyPair[n].tab1 = col1; m_KeyPair[n].tab2 = col2; }
    void pushKey(int col1, int col2)
    {
        m_KeyPair[m_nKeyPairCounter].tab1 = col1;
        m_KeyPair[m_nKeyPairCounter].tab2 = col2;
        m_nKeyPairCounter++;
    }
    void deleteAllKeys()
    {
        for (int i = 0; i < m_nKeyPairCounter; i++) { m_KeyPair[i].tab1 = 0; m_KeyPair[i].tab2 = 0; }
        m_nKeyPairCounter = 0;
    }
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
    void deleteKeyAt(int n)
    {
        for (int i = n; i < m_nKeyPairCounter; i++)
        {
            m_KeyPair[i].tab1 = m_KeyPair[i + 1].tab1;
            m_KeyPair[i].tab2 = m_KeyPair[i + 1].tab2;
        }
        m_nKeyPairCounter--;
    }
    void insertKeyAt(int n, int col1, int col2)
    {
        for (int i = m_nKeyPairCounter; i > n; i--)
        {
            m_KeyPair[i].tab1 = m_KeyPair[i - 1].tab1;
            m_KeyPair[i].tab2 = m_KeyPair[i - 1].tab2;
        }
        m_nKeyPairCounter++;
    }
    bool isThisAKey(int table, int column) const
    {
        for (int i = 0; i < m_nKeyPairCounter; i++)
        {
            if (table == 1 && m_KeyPair[i].tab1 == column) return true;
            if (table == 2 && m_KeyPair[i].tab2 == column) return true;
        }
        return false;
    }
    bool areThereAnyKeys() const { return m_nKeyPairCounter > 0; }
    int  getKeyPairCounter() const { return m_nKeyPairCounter; }

    // --- Data accessors ---
    CString getKeyStr1(int row) const { return m_pszKeyArr11[row]; }
    CString getKeyStr2(int row) const { return m_pszKeyArr21[row]; }
    bool    isKeyMissing1(int row) const { return m_pbKeyMissing1[row]; }
    bool    isKeyMissing2(int row) const { return m_pbKeyMissing2[row]; }
    bool    isEmptyCol1(int col)   const { return m_pbEmptyClms1[col]; }
    bool    isEmptyCol2(int col)   const { return m_pbEmptyClms2[col]; }
    char    getMainChar1(int row, int col) const { return m_pchMainArr1[(row - 1) * m_Table1.NumberOfColumns + col]; }
    char    getMainChar2(int row, int col) const { return m_pchMainArr2[(row - 1) * m_Table2.NumberOfColumns + col]; }
    CMap<CString, LPCTSTR, long, long>& getMap1() { return m_Map1; }
    CMap<CString, LPCTSTR, long, long>& getMap2() { return m_Map2; }

    NotUniqueKeys m_NotUniqueKeys1;
    NotUniqueKeys m_NotUniqueKeys2;
    bool          m_bUseIndexes = false;

private:
    // --- Internal helpers ---
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
