#pragma once

// Key candidate columns for one table
struct PossibleKeys {
	int k[256];
};

// RGB color entry for the difference-highlight palette
struct Palette {
	int red;
	int green;
	int blue;
};

// Descriptor of one Excel table (sheet)
struct Table {
	int    WorkSheetNumber;
	long   MaxNumberOfRows;
	long   MaxNumberOfCols;
	long   NumberOfRows;
	int    FirstRowWithData;
	int    RowWithNames;
	int    NumberOfColumns;
	CString Columns[256];
	bool   keys[256];
	int    keysCnt;
};

// Top-left corner of the visible matrix region (in cell units)
struct VisTopLeft {
	int top;
	int left;
};

// One key-column pair (one column from each table)
struct KeyPair {
	int tab1;
	int tab2;
};

// Similarity score between two columns across tables
struct SimilaritiesAcrossTables {
	int  clm1;
	int  clm2;
	long similarity;
	int  similarityOrder;
	int  pureSim;
};

// Matrix cell coordinates (column x, row y)
struct ChosenCell {
	int x;
	int y;
};

// Size of the client drawing area in pixels
struct Clnt {
	int w;
	int h;
};

// Best key combination found by the key-suggestion algorithm
struct BestKeyComb {
	int  pk1;
	int  pk2;
	int  rating;
	long cnt;
};

// Information about the first duplicate key found (used for user error messages)
struct NotUniqueKeys {
	long    firstRow;
	long    secondRow;
	CString keyString;
};
