// Copyright (c) 2014-2018 Thomas Fussell
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, WRISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE
//
// @license: http://www.opensource.org/licenses/mit-license.php
// @author: see AUTHORS file

#pragma once
#include "CoreMinimal.h"
#include <string>
#include <algorithm>
#include <sstream>
#include <sstream>

//#if !defined(_WIN32)&&!defined(__APPLE__)
//
//#include <sstream>
//
//
//#endif // !_WIN32

namespace xlnt_custom
{
	std::string to_string(uint32 value);

	

	template <typename T>
	std::string to_string(T value)
	{
		std::ostringstream os;
		os << value;
		return os.str();
	}

	inline int stoi(std::string s)
	{
		return atoi(s.c_str());
	}
	inline double stod(std::string s)
	{
		return atof(s.c_str());
	}

	inline int stoi16(std::string s)
	{
		int i = 0;
#ifdef _WIN32
		sscanf_s(s.c_str(), "%x", &i);
#else
		sscanf(s.c_str(), "%x", &i);
#endif
		return i;
	}

	inline int stoul(std::string s)
	{
		return atoi(s.c_str());
	}
	inline int stoull(std::string s)
	{
		return atoi(s.c_str());
	}

	inline int round(double v)
	{
		return ::round(v);
	}

	inline bool string_contains(const std::string& str, char c)
	{
		for (auto i : str)
		{
			if (i == c)
			{
				return true;
			}
		}
		return false;
	}

	template<typename TI, typename TV>
	inline TI find(TI first, TI last, const TV& val)
	{
		for (; first != last; ++first)
		{
			if (*first == val)
			{
				break;
			}
		}
		return first;
	}

	template <class _InIt, class _Diff>
	inline void advance(_InIt& _Where, _Diff _Off)
	{ // increment iterator by offset
		for (; 0 < _Off; --_Off) {
			++_Where;
		}
	}

	template <class _InIt, class _Pr>
	inline _InIt find_if(_InIt _First, const _InIt _Last, _Pr _Pred) { // find first satisfying _Pred
		for (; _First != _Last; ++_First) {
			if (_Pred(*_First)) {
				break;
			}
		}

		return _First;
	}

	template <class T>
	inline T min(const T& a, const T b)
	{
		return a > b ? b : a;
	}

	template <class T>
	inline T max(const T& a, const T b)
	{
		return a > b ? a : b;
	}

	template <class _FwdIt, class _Ty>
	inline void fill(_FwdIt _First, _FwdIt _Last, const _Ty& _Val) { // copy _Val through [_First, _Last)
		for (; _First != _Last; ++_First) {
			*_First = _Val;
		}
	}

	template<class InputIterator, class OutputIterator>
	inline OutputIterator copy(InputIterator first, InputIterator last, OutputIterator result)
	{
		while (first != last) {
			*result = *first;
			++result; ++first;
		}
		return result;
	}

	template<typename T>
	inline T fabs(T x) { return ::fabs(double(x)); }
}