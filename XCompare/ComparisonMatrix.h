#pragma once
#include <vector>

// Encapsulates the 2D result matrix produced by the comparison algorithm.
// The matrix stores, for each (col2, col1) cell, how many matching rows
// were found between the two tables. Additionally it tracks which cells
// the user has manually marked for follow-up inspection.
//
// Coordinate convention (inherited from original code):
//   x  = column index from table 2  (1-based)
//   y  = row index    from table 1  (1-based)
// Both axes are therefore 1-based; index 0 is unused.
class ComparisonMatrix
{
public:
    ComparisonMatrix() : m_width(0), m_height(0) {}

    // Resize the matrix to (x+1)*(y+1) and zero-initialise all cells.
    // x = table-2 column count, y = table-1 column count.
    void clear(int x, int y)
    {
        m_width = x;
        m_height = y;
        int size = (x + 1) * (y + 1);
        m_values.assign(size, 0);
        m_marked.assign(size, false);
    }

    // Increment the match counter at (x, y) by one.
    void increment(int x, int y) { m_values[index(x, y)] += 1; }

    // Return the match counter at (x, y).
    [[nodiscard]] int get(int x, int y) const { return m_values[index(x, y)]; }

    // Return whether (x, y) is user-marked.
    [[nodiscard]] bool isMarked(int x, int y) const { return m_marked[index(x, y)]; }

    // Mark (x, y) as user-selected.
    void setMarked(int x, int y) { m_marked[index(x, y)] = true; }

    // Clear the marked flag at (x, y).
    void clearMarked(int x, int y) { m_marked[index(x, y)] = false; }

private:
    int linearIndex(int x, int y) const { return (y - 1) * m_width + x; }
    int index(int x, int y) const { return linearIndex(x, y); }

    int m_width;                // number of table-2 columns
    int m_height;               // number of table-1 columns
    std::vector<int> m_values;  // match counters
    std::vector<bool> m_marked; // user-marked flags
};
