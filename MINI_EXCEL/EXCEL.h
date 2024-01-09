#pragma once
#include <iostream>
#include <windows.h>
#include <conio.h>
#include <fstream>
#include <string>
#include <sstream>
#include <tuple>
#include <vector>
#include <regex>

using namespace std;

enum COLOR
{
	BLACK = 0,
	WHITE = 7,
	GREEN = 2,
	BLUE = 9
};
enum DATATYPE
{
	Int,
	Float,
	String
};
bool isInteger(string s)
{
	istringstream iss(s);
	int n;
	iss >> noskipws >> n;
	return iss.eof() && !iss.fail();
}
bool isFloat(string s)
{
	istringstream iss(s);
	float f;
	iss >> noskipws >> f;
	return iss.eof() && !iss.fail();
}
string Spaces(string s)
{
	regex pattern("\\s+$");
	return regex_replace(s, pattern, "");
}
string Arr[10000000];
int ArrCount = 0;
class Excel
{
public:
	class Cell
	{
		int X;
		int Y;
		COLOR color;
		DATATYPE type;
		string DATA = "    ";
	public:
		Cell()
		{
			X = Y = 0;
			DATA = "";
			color = WHITE;
		}
		Cell(int x, int y, string data)
		{
			X = x;
			Y = y;
			DATA = data;
			color = WHITE;
		}
		string getData() { return DATA; }
		int getX() { return X; }
		int getY() { return Y; }
		void setX(int x) { X = x; }
		void setY(int y) { Y = y; }
		void select() { color = GREEN; }
		void deSelect() { color = WHITE; }
		void setData(string data)
		{
			if (isInteger(data))
				setDataType(1);
			else if (isFloat(data))
				setDataType(2);
			else
				setDataType(3);
			string s = "";
			for (int i = 0; i < 4 && i < data.length(); i++)
			{
				s += data[i];
			}
			if (data.length() < 4)
			{
				for (int i = 0; i < 4 - data.length(); i++)
				{
					s += " ";
				}
			}
			this->DATA = s;
		}
		void setDataType(int key)
		{
			if (key == 1)
				type = Int;
			else if (key == 2)
				type = Float;
			else
				type = String;
		}
		DATATYPE GetDataType() { return type; }

		friend class Excel;
	};
	class Node
	{
		Cell* DATA;
		Node* UP;
		Node* DOWN;
		Node* RIGHT;
		Node* LEFT;
	public:
		Node(Node* up = nullptr, Node* down = nullptr, Node* right = nullptr, Node* left = nullptr)
		{
			DATA = new Cell();
			UP = up;
			DOWN = down;
			RIGHT = right;
			LEFT = left;
		}
		Node(Cell* data)
		{
			DATA = data;
			UP = nullptr;
			DOWN = nullptr;
			RIGHT = nullptr;
			LEFT = nullptr;
		}
		friend class Excel;
	};
	class iterator
	{
		Node* iter;
	public:
		iterator()
		{
			iter = nullptr;
		}
		iterator(Node* n)
		{
			iter = n;
		}
		iterator operator++()
		{
			if (iter->DOWN != nullptr)
				iter = iter->DOWN;
			return *this;
		}
		iterator operator--()
		{
			if (iter->UP != nullptr)
				iter = iter->UP;
			return *this;
		}
		iterator operator++(int)
		{
			if (iter->RIGHT != nullptr)
				iter = iter->RIGHT;
			return *this;
		}
		iterator operator--(int)
		{
			if (iter->LEFT != nullptr)
				iter = iter->LEFT;
			return *this;
		}
		bool operator ==(iterator i)
		{
			return (iter == i.iter);
		}
		bool operator!=(iterator i)
		{
			return (iter != i.iter);
		}
		Cell*& operator*()
		{
			return iter->DATA;
		}

		friend class Excel;
	};
	Excel()
	{
		HEAD = nullptr;
		CURRENT = nullptr;

		rowSize = colSize = 5;
		curRow = curCol = 0;

		HEAD = newRow();
		CURRENT = HEAD;
		for (int i = 0;i < rowSize - 1;i++)
		{
			InsertRowBelow();
			rowSize--;
			CURRENT = CURRENT->DOWN;
		}
		CURRENT = HEAD;

		PrintGrid();
		PrintData();
	}
	Node* getCell(int x, int y)
	{
		Node* temp = topLeft();
		for (int i = 0; i < x; i++)
		{
			temp = temp->DOWN;
		}
		for (int i = 0; i < y; i++)
		{
			temp = temp->RIGHT;
		}
		return temp;
	}
	Node* topRight()
	{
		Node* right = CURRENT;
		while (right->UP)
		{
			right = right->UP;
		}
		while (right->RIGHT)
		{
			right = right->RIGHT;
		}
		return right;
	}
	Node* bottomRight()
	{
		Node* right = CURRENT;
		while (right->DOWN)
		{
			right = right->DOWN;
		}
		while (right->RIGHT)
		{
			right = right->RIGHT;
		}
		return right;
	}
	Node* topLeft()
	{
		Node* left = CURRENT;
		while (left->UP)
		{
			left = left->UP;
		}
		while (left->LEFT)
		{
			left = left->LEFT;
		}
		return left;
	}
	Node* bottomLeft()
	{
		Node* left = CURRENT;
		while (left->DOWN)
		{
			left = left->DOWN;
		}
		while (left->LEFT)
		{
			left = left->LEFT;
		}
		return left;
	}
	void moveLeft()
	{
		if (CURRENT->UP == nullptr)
		{
			InsertColLeft();
			PrintRow();
			curRow++;
		}
		PrintCell(curRow, curCol, WHITE);
		CURRENT = CURRENT->UP;
		curRow--;
	}
	void moveUp()
	{
		if (CURRENT->LEFT == nullptr)
		{
			InsertRowAbove();
			PrintColumn();
			curCol++;
		}
		PrintCell(curRow, curCol, WHITE);
		CURRENT = CURRENT->LEFT;
		curCol--;
	}
	void moveRight()
	{
		if (CURRENT->DOWN == nullptr)
		{
			InsertRowBelow();
			PrintRow();
		}
		PrintCell(curRow, curCol, WHITE);
		CURRENT = CURRENT->DOWN;
		curRow++;
	}
	void moveDown()
	{
		if (CURRENT->RIGHT == nullptr)
		{
			InsertColRight();
			PrintColumn();
		}
		PrintCell(curRow, curCol, WHITE);
		CURRENT = CURRENT->RIGHT;
		curCol++;
	}
	Node* newRow()
	{
		Node* temp = new Node();
		Node* curr = temp;
		for (int i = 0; i < colSize - 1; i++)
		{
			Node* temp2 = new Node();
			temp->RIGHT = temp2;
			temp2->LEFT = temp;
			temp = temp2;
		}
		return curr;
	}
	Node* newCol()
	{
		Node* temp = new Node();
		Node* curr = temp;
		for (int i = 0; i < rowSize - 1; i++)
		{
			Node* temp2 = new Node();
			temp->DOWN = temp2;
			temp2->UP = temp;
			temp = temp2;
		}
		return curr;
	}
	void InsertColRight()
	{
		Node* temp = newCol();
		Node* temp2 = CURRENT;
		while (temp2->UP != nullptr)
		{
			temp2 = temp2->UP;
		}
		if (temp2->RIGHT == nullptr)
		{
			while (temp2 != nullptr)
			{
				temp2->RIGHT = temp;
				temp->LEFT = temp2;
				temp2 = temp2->DOWN;
				temp = temp->DOWN;
			}
		}
		else
		{
			while (temp2 != nullptr)
			{
				temp->RIGHT = temp2->RIGHT;
				temp2->RIGHT = temp;
				temp->LEFT = temp2;
				temp->RIGHT->LEFT = temp;
				temp2 = temp2->DOWN;
				temp = temp->DOWN;
			}
		}
		colSize++;
	}
	void InsertColLeft()
	{
		Node* temp = newRow();
		Node* temp2 = CURRENT;
		while (temp2->LEFT != nullptr)
		{
			temp2 = temp2->LEFT;
		}
		if (temp2 == HEAD)
		{
			HEAD = temp;
		}
		if (temp2->UP == nullptr)
		{
			while (temp2 != nullptr)
			{
				temp2->UP = temp;
				temp->DOWN = temp2;
				temp = temp->RIGHT;
				temp2 = temp2->RIGHT;
			}
		}
		else
		{
			while (temp2 != nullptr)
			{
				temp->UP = temp2->UP;
				temp2->UP = temp;
				temp->DOWN = temp2;
				temp->UP->DOWN = temp;
				temp = temp->RIGHT;
				temp2 = temp2->RIGHT;
			}
		}
		rowSize++;
		PrintGrid();
		PrintData();
	}
	void InsertRowAbove()
	{
		Node* temp = newCol();
		Node* temp2 = CURRENT;
		while (temp2->UP != nullptr)
		{
			temp2 = temp2->UP;
		}
		if (temp2 == HEAD)
		{
			HEAD = temp;
		}
		if (temp2->LEFT == nullptr)
		{
			while (temp2 != nullptr)
			{
				temp2->LEFT = temp;
				temp->RIGHT = temp2;
				temp2 = temp2->DOWN;
				temp = temp->DOWN;
			}
		}
		else
		{
			while (temp2 != nullptr)
			{
				temp->LEFT = temp2->LEFT;
				temp2->LEFT = temp;
				temp->RIGHT = temp2;
				temp->LEFT->RIGHT = temp;
				temp2 = temp2->DOWN;
				temp = temp->DOWN;
			}
		}
		colSize++;
		PrintGrid();
		PrintData();
	}
	void InsertRowBelow()
	{
		Node* temp = newRow();
		Node* temp2 = CURRENT;
		while (temp2->LEFT != nullptr)
		{
			temp2 = temp2->LEFT;
		}
		if (temp2->DOWN == nullptr)
		{
			while (temp2 != nullptr)
			{
				temp2->DOWN = temp;
				temp->UP = temp2;
				temp = temp->RIGHT;
				temp2 = temp2->RIGHT;
			}
		}
		else
		{
			while (temp2 != nullptr)
			{
				temp->DOWN = temp2->DOWN;
				temp2->DOWN = temp;
				temp->UP = temp2;
				temp->DOWN->UP = temp;
				temp = temp->RIGHT;
				temp2 = temp2->RIGHT;
			}
		}
		rowSize++;
	}
	void InsertCellByRightShift()
	{
		Node* temp = CURRENT;
		while (CURRENT->RIGHT != nullptr)
		{
			CURRENT = CURRENT->RIGHT;
		}
		InsertColRight();
		CURRENT = CURRENT->RIGHT;
		while (CURRENT != temp)
		{
			CURRENT->DATA = CURRENT->LEFT->DATA;
			CURRENT = CURRENT->LEFT;
		}
		CURRENT->DATA->setData("");
	}
	void InsertCellByDownShift()
	{
		Node* temp = CURRENT;
		while (CURRENT->DOWN != nullptr)
		{
			CURRENT = CURRENT->DOWN;
		}
		InsertRowBelow();
		CURRENT = CURRENT->DOWN;
		while (CURRENT != temp)
		{
			CURRENT->DATA = CURRENT->UP->DATA;
			CURRENT = CURRENT->UP;
		}
		CURRENT->DATA->setData("");
	}
	void addRow()
	{
		Node* newNode = bottomLeft();
		while (newNode)
		{
			Node* node = new Node();
			newNode->DOWN = node;
			newNode->DOWN->UP = newNode;
			newNode = newNode->RIGHT;
		}
		newNode = bottomLeft();
		while (newNode->UP->RIGHT)
		{
			newNode->RIGHT = newNode->UP->RIGHT->DOWN;
			newNode->RIGHT->LEFT = newNode;
			newNode = newNode->RIGHT;
		}
		rowSize++;
	}
	void addColumn()
	{
		Node* newNode = topRight();
		while (newNode)
		{
			Node* node = new Node();
			newNode->RIGHT = node;
			newNode->RIGHT->LEFT = newNode;
			newNode = newNode->DOWN;
		}
		newNode = topRight();
		while (newNode->LEFT->DOWN)
		{
			newNode->DOWN = newNode->LEFT->DOWN->RIGHT;
			newNode->DOWN->UP = newNode;
			newNode = newNode->DOWN;
		}
		colSize++;
	}
	void color(int k)
	{
		HANDLE hConsole = GetStdHandle(STD_OUTPUT_HANDLE);
		SetConsoleTextAttribute(hConsole, k);
		if (k > 255)
		{
			k = 1;
		}
	}
	void gotoXY(int x, int y)
	{
		COORD coordinates;
		coordinates.X = x;
		coordinates.Y = y;
		SetConsoleCursorPosition(GetStdHandle(STD_OUTPUT_HANDLE), coordinates);
	}
	void PrintCell(int row, int col, int colour)
	{
		color(colour);
		char border = 254;

		int x = row * width;
		int y = col * height;

		gotoXY(x, y);
		for (int i = 0; i < width; i++)
		{
			cout << border;
		}
		gotoXY(x, y + height);
		for (int i = 0; i < width; i++)
		{
			cout << border;
		}
		for (int i = 0; i < height; i++)
		{
			gotoXY(x, y + i);
			cout << border;
		}
		for (int i = 0; i <= height; i++)
		{
			gotoXY(x + width, y + i);
			cout << border;
		}
		gotoXY((x + width / 2) - 3, y + height / 2);
		cout << "";
	}
	void PrintCellData(int row, int col, Node* d, int colour)
	{
		color(colour);
		gotoXY((width * row) + width / 2, (col * height) + height / 2);
		cout << d->DATA->getData();
	}
	void PrintGrid()
	{
		system("cls");
		for (int i = 0;i < rowSize;i++)
		{
			for (int j = 0;j < colSize;j++)
			{
				PrintCell(i, j, WHITE);
			}
		}
	}
	void PrintData()
	{
		Node* temp = HEAD;
		for (int ri = 0; ri < rowSize; ri++)
		{
			Node* temp2 = temp;
			for (int ci = 0; ci < colSize; ci++)
			{
				PrintCellData(ri, ci, temp, WHITE);
				temp = temp->RIGHT;
			}

			temp = temp2->DOWN;
		}
	}
	void PrintColumn()
	{
		for (int i = 0, j = curCol + 1;i < rowSize;i++)
		{
			PrintCell(i, j, WHITE);
		}
	}
	void PrintRow()
	{
		for (int j = 0, i = curRow + 1;j < colSize;j++)
		{
			PrintCell(i, j, WHITE);
		}
	}
	void inputData()
	{
		string input = "";
		color(WHITE);
		cin >> input;
		if (input.length() > 4)
		{
			input = "";
			PrintGrid();
			PrintData();
		}
		else
		{
			CURRENT->DATA->setData(input);
			input = "";
		}
	}
	void SaveData()
	{
		Node* temp = HEAD;
		ofstream fout("SaveExcel.txt");

		fout << rowSize << endl;
		fout << colSize << endl;

		for (int i = 0; i < rowSize; i++)
		{
			Node* temp2 = temp;
			for (int j = 0; j < colSize; j++)
			{
				if (temp->DATA->getData() == "")
				{
					fout << "NULL"
						<< " ";
				}
				else
				{
					fout << temp->DATA->getData() << " ";
				}

				temp = temp->RIGHT;
			}
			fout << endl;
			temp = temp2->DOWN;
		}
	}
	void LoadData()
	{
		system("cls");
		ifstream fin("SaveExcel.txt");
		fin >> rowSize;
		fin >> colSize;

		HEAD = nullptr;
		CURRENT = nullptr;
		curRow = curCol = 0;

		HEAD = newRow();
		CURRENT = HEAD;
		for (int i = 0; i < rowSize - 1; i++)
		{
			InsertRowBelow();
			rowSize--;
			CURRENT = CURRENT->DOWN;
		}

		string data;
		CURRENT = HEAD;
		Node* temp = CURRENT;
		for (int ri = 0; ri < rowSize; ri++)
		{
			Node* temp2 = temp;
			for (int ci = 0; ci < colSize; ci++)
			{
				fin >> data;
				if (data == "NULL")
					temp->DATA->setData("");
				else
					temp->DATA->setData(data);

				temp = temp->RIGHT;
			}

			temp = temp2->DOWN;
		}

		PrintGrid();
		PrintData();
	}
	void ClearRow()
	{
		/*Node* temp1 = CURRENT;
		while (temp1->UP)
		{
			temp1 = temp1->UP;
		}
		while (temp1->DOWN)
		{
			temp1->DATA->setData(" ");
			temp1 = temp1->DOWN;
		}
		PrintGrid();
		PrintData();*/
	}
	void ClearColumn()
	{
		Node* temp1 = CURRENT;
		while (temp1->LEFT)
		{
			temp1 = temp1->LEFT;
		}
		while (temp1->RIGHT)
		{
			temp1->DATA->setData(" ");
			temp1 = temp1->RIGHT;
		}
		PrintGrid();
		PrintData();
	}
	void DeleteCellUp()
	{
		Node* currentCell = CURRENT;
		if (currentCell->LEFT != nullptr)
		{
			currentCell->DATA->setData("");
			Node* temp = currentCell->RIGHT;
			while (temp != nullptr)
			{
				currentCell->DATA->setData(temp->DATA->getData());
				currentCell = temp;
				currentCell->DATA->setData("");
				temp = temp->RIGHT;
			}
		}
	}
	void DeleteCellLeft()
	{
		Node* currentCell = CURRENT;
		if (currentCell->UP != nullptr)
		{
			currentCell->DATA->setData("");
			Node* temp = currentCell->DOWN;
			while (temp != nullptr)
			{
				currentCell->DATA->setData(temp->DATA->getData());
				currentCell = temp;
				currentCell->DATA->setData("");
				temp = temp->DOWN;
			}
		}
	}
	void DeleteRow()
	{
		//rowSize--;
		/*Node* curr = CURRENT;
		while (curr->LEFT)
		{
			curr = curr->LEFT;
		}
		Node* previous = nullptr;
		while (curr)
		{
			Node* top = curr->LEFT;
			if (top)
			{
				top->RIGHT = curr->RIGHT;
			}

			if (curr->RIGHT)
			{
				curr->RIGHT->LEFT = top;
			}

			if (curr->DOWN)
			{
				curr->DOWN->UP = curr->UP;
			}
			if (curr->UP)
			{
				curr->UP->DOWN = curr->DOWN;
			}
			previous = curr;
			curr = curr->DOWN;
			delete previous;
		}
		if (CURRENT->UP != nullptr)
			CURRENT = CURRENT->UP;
		else
			CURRENT = CURRENT->DOWN;*/
		PrintGrid();
		PrintData();
	}
	void DeleteColumn()
	{
		//colSize--;
		/*Node* curr = CURRENT;
		while (curr->LEFT)
		{
			curr = curr->LEFT;
		}
		Node* previous = nullptr;
		while (curr)
		{
			Node* top = curr->LEFT;
			if (top)
			{
				top->RIGHT = curr->RIGHT;
			}

			if (curr->RIGHT)
			{
				curr->RIGHT->LEFT = top;
			}

			if (curr->DOWN)
			{
				curr->DOWN->UP = curr->UP;
			}
			if (curr->UP)
			{
				curr->UP->DOWN = curr->DOWN;
			}
			previous = curr;
			curr = curr->DOWN;
			delete previous;
		}
		if (CURRENT->UP != nullptr)
			CURRENT = CURRENT->UP;
		else
			CURRENT = CURRENT->DOWN;*/
		PrintGrid();
		PrintData();
	}
	int SumCalculator(Node* startingCell, Node* endingCell)
	{
		int x = startingCell->DATA->getX();
		int x1 = endingCell->DATA->getX();
		int y = startingCell->DATA->getY();
		int y1 = endingCell->DATA->getY();
		DATATYPE d;
		int cal = 0;
		if (x == x1)
		{
			while (startingCell != endingCell->RIGHT)
			{
				d = startingCell->DATA->GetDataType();
				if (d == Int || d == Float)
				{
					cal += stoi(Spaces(startingCell->DATA->getData()));
					startingCell = startingCell->RIGHT;
				}
				else
				{
					startingCell = startingCell->RIGHT;
				}
			}
		}
		else
		{
			while (startingCell != endingCell->DOWN)
			{
				d = startingCell->DATA->GetDataType();
				if (d == Int || d == Float)
				{
					cal += stoi(Spaces(startingCell->DATA->getData()));
					startingCell = startingCell->DOWN;
				}
				else
				{
					startingCell = startingCell->DOWN;
				}
			}
		}
		return cal;
	}
	void sumOfRange()
	{
		int x1, x2, y1, y2;
		vector<int>coords = getRange();
		x1 = coords[0];
		y1 = coords[1];
		x2 = coords[2];
		y2 = coords[3];
		if (x1 == x2 || y1 == y2)
		{
			int val = SumCalculator(getCell(x1, y1), getCell(x2, y2));
			CURRENT->DATA->setData(to_string(val));
		}
		else
		{
			cout << "Enter cells that are in same row or same column" << endl;
			cin.get();
		}
		PrintGrid();
		PrintData();
	}
	float AverageCalculator(Node* startingCell, Node* endingCell)
	{
		int x = startingCell->DATA->getX();
		int x1 = endingCell->DATA->getX();
		int y = startingCell->DATA->getY();
		int y1 = endingCell->DATA->getY();
		DATATYPE d;
		float cal = 0;
		float counter = 0;
		if (x == x1)
		{
			while (startingCell != endingCell->RIGHT)
			{
				d = startingCell->DATA->GetDataType();
				if (d == Int || d == Float)
				{
					cal += stoi(Spaces(startingCell->DATA->getData()));
					startingCell = startingCell->RIGHT;
					counter++;
				}
				else
				{
					startingCell = startingCell->RIGHT;
				}
			}
		}
		else
		{
			while (startingCell != endingCell->DOWN)
			{
				d = startingCell->DATA->GetDataType();
				if (d == Int || d == Float)
				{
					cal += stoi(Spaces(startingCell->DATA->getData()));
					startingCell = startingCell->DOWN;
					counter++;
				}
				else
				{
					startingCell = startingCell->DOWN;
				}
			}
		}
		return cal / counter;
	}
	void avgOfRange()
	{
		int x1, x2, y1, y2;
		vector<int>coords = getRange();
		x1 = coords[0];
		y1 = coords[1];
		x2 = coords[2];
		y2 = coords[3];
		if (x1 == x2 || y1 == y2)
		{
			float val = AverageCalculator(getCell(x1, y1), getCell(x2, y2));
			CURRENT->DATA->setData(to_string(val));
		}
		else
		{
			cout << "Enter cells that are in same row or same column" << endl;
			cin.get();
		}
		PrintGrid();
		PrintData();
	}
	void CutRange(Node* startingCell, Node* endingCell)
	{
		ArrCount = 0;
		int x = startingCell->DATA->getX();
		int x1 = endingCell->DATA->getX();
		int y = startingCell->DATA->getY();
		int y1 = endingCell->DATA->getY();
		if (x == x1)
		{
			while (startingCell != endingCell->DOWN)
			{
				Arr[ArrCount] = startingCell->DATA->getData();
				startingCell = startingCell->DOWN;
				ArrCount++;
			}
		}
		else
		{
			while (startingCell != endingCell->RIGHT)
			{
				Arr[ArrCount] = startingCell->DATA->getData();
				startingCell = startingCell->RIGHT;
				ArrCount++;
			}
		}
	}
	void Cut()
	{
		int x1, x2, y1, y2;
		vector<int>coords = getRange();
		x1 = coords[0];
		y1 = coords[1];
		x2 = coords[2];
		y2 = coords[3];
		if (x1 == x2 || y1 == y2)
		{
			CutRange(getCell(x1, y1), getCell(x2, y2));
			cout << "Copied";
		}
		else
		{
			cout << "Enter cells that are in same row or same column" << endl;
			cin.get();
		}
		PrintGrid();
		PrintData();
	}
	void CopyRange(Node* startingCell, Node* endingCell)
	{
		ArrCount = 0;
		int x = startingCell->DATA->getX();
		int x1 = endingCell->DATA->getX();
		int y = startingCell->DATA->getY();
		int y1 = endingCell->DATA->getY();
		if (x == x1)
		{
			while (startingCell != endingCell->DOWN)
			{
				Arr[ArrCount] = startingCell->DATA->getData();
				startingCell = startingCell->DOWN;
				ArrCount++;
			}
		}
		else
		{
			while (startingCell != endingCell->RIGHT)
			{
				Arr[ArrCount] = startingCell->DATA->getData();
				startingCell = startingCell->RIGHT;
				ArrCount++;
			}
		}
	}
	void Copy()
	{
		int x1, x2, y1, y2;
		vector<int>coords = getRange();
		x1 = coords[0];
		y1 = coords[1];
		x2 = coords[2];
		y2 = coords[3];
		if (x1 == x2 || y1 == y2)
		{
			CopyRange(getCell(x1, y1), getCell(x2, y2));
			cout << "Cut";
		}
		else
		{
			cout << "Enter cells that are in same row or same column" << endl;
			cin.get();
		}
		PrintGrid();
		PrintData();
	}
	void PasteRange(string val)
	{
		Node* temp2 = CURRENT;
		if (val == "Row" || val == "row")
		{
			for (int i = 0; i < ArrCount; i++)
			{
				temp2->DATA->setData(Arr[i]);
				if (temp2->DOWN == nullptr)
				{
					addColumn();
				}
				temp2 = temp2->DOWN;
			}
		}
		else
		{
			for (int i = 0; i < ArrCount; i++)
			{
				temp2->DATA->setData(Arr[i]);
				if (temp2->RIGHT == nullptr)
				{
					addRow();
				}
				temp2 = temp2->RIGHT;
			}
		}
	}
	void Paste()
	{
		string value;
		cout << endl;
		cout << "Enter where you want to paste data i.e in row or column: ";
		cin >> value;
		if (value == "Row" || value == "row" || value == "Column" || value == "column")
		{
			PasteRange(value);
		}
		else
		{
			cout << "Please enter correct string" << endl;
			cin.get();
		}
		PrintGrid();
		PrintData();
	}
	vector<int> getRange()
	{
		int x1 = 0, x2 = 0, y1 = 0, y2 = 0;
		cout << endl;
		cout << "Enter Row number of starting cell : ";
		cin >> x1;
		cout << "Enter Column number of starting cell : ";
		cin >> y1;
		cout << "Enter Row number of ending cell : ";
		cin >> x2;
		cout << "Enter Column Number of ending cell : ";
		cin >> y2;
		return vector<int>{x1, y1, x2, y2};
		// you can also use tuple as explained below;
		//tuple<tuple<int, int>, tuple<int, int>>coords = make_tuple(make_tuple(x1, y1), make_tuple(x2, y2));

	}
	void Menu()
	{
		system("cls");
		cout << "------------- MINI EXCEL -------------" << endl << endl;

		cout << "Press F1 for Help" << endl;

		cout << "Press esc to EXIT" << endl;

		cout << "Press F10 to LOAD data from file" << endl;
		cout << "Press F12 to SAVE data in file" << endl;

		cout << "Press A || a to EXTEND row above" << endl;
		cout << "Press S || s to EXTEND row below " << endl;
		cout << "Press D || d to EXTEND column left" << endl;
		cout << "Press F || f to EXTEND column right" << endl;

		cout << "Press H || h to DELETE row" << endl;
		cout << "Press J || j to DELETE column" << endl << endl;
		cout << "Press K || k to CLEAR row" << endl;
		cout << "Press L || l to CLEAR column" << endl;

		cout << "Press Q || q to INSERT cell right" << endl;
		cout << "Press W || w to INSERT cell down" << endl;
		cout << "Press E || e to DELETE cell left " << endl;
		cout << "Press R || r to DELETE cell up" << endl;

		cout << "Press O || o to SUM of Range" << endl;
		cout << "Press I || i to AVERAGE of range" << endl;

		cout << "Press X || x to CUT" << endl;
		cout << "Press C || c to COPY" << endl;
		cout << "Press V || v to PASTE" << endl;


		cout << "Press ENTER to continue .....";
		cin.get();
		Run();
	}
	void Run()
	{
		PrintGrid();
		PrintCell(curRow, curCol, GREEN);
		while (true)
		{
			if (GetAsyncKeyState(VK_F1)) { Menu(); }
			if (GetAsyncKeyState(VK_SPACE)) { inputData(); }
			if (GetAsyncKeyState(VK_RIGHT)) { moveRight(); }
			if (GetAsyncKeyState(VK_LEFT)) { moveLeft(); }
			if (GetAsyncKeyState(VK_UP)) { moveUp(); }
			if (GetAsyncKeyState(VK_DOWN)) { moveDown(); }
			if (GetAsyncKeyState(VK_F10)) { LoadData(); }
			if (GetAsyncKeyState(VK_F12)) { SaveData(); }
			if (GetAsyncKeyState(VK_ESCAPE)) { exit(0); }

			char code = _getch();
			if (code == 'A' || code == 'a') { InsertRowAbove(); }
			else if (code == 'S' || code == 's') { InsertRowBelow(); PrintGrid();PrintData(); }
			else if (code == 'D' || code == 'd') { InsertColLeft(); }
			else if (code == 'F' || code == 'f') { InsertColRight();PrintGrid();PrintData(); }
			else if (code == 'Q' || code == 'q') { InsertCellByRightShift();PrintGrid();PrintData(); }
			else if (code == 'W' || code == 'w') { InsertCellByDownShift(); PrintGrid();PrintData(); }
			else if (code == 'E' || code == 'e') { DeleteCellLeft(); PrintGrid();PrintData(); }
			else if (code == 'R' || code == 'r') { DeleteCellUp(); PrintGrid();PrintData(); }
			else if (code == 'H' || code == 'h') { DeleteRow(); }
			else if (code == 'J' || code == 'j') { DeleteColumn(); }
			else if (code == 'K' || code == 'k') { ClearRow(); }
			else if (code == 'L' || code == 'l') { ClearColumn(); }
			else if (code == 'O' || code == 'o') { sumOfRange(); }
			else if (code == 'I' || code == 'i') { avgOfRange(); }
			else if (code == 'X' || code == 'x') { Cut(); }
			else if (code == 'C' || code == 'c') { Copy(); }
			else if (code == 'V' || code == 'v') { Paste(); }
			else { code = ' '; }
			code = ' ';
			gotoXY(curRow, curCol);
			PrintCell(curRow, curCol, GREEN);
			PrintCellData(curRow, curCol, CURRENT, WHITE);
			Sleep(100);
		}
	}
private:
	Node* CURRENT;
	Node* HEAD;
	int rowSize, colSize;
	int curRow, curCol;
	int width = 8;
	int height = 4;
};