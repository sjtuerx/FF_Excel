// file      : xml/value-traits.cxx
// copyright : Copyright (c) 2013-2017 Code Synthesis Tools CC
// license   : MIT; see accompanying LICENSE file

#include "parser.h"

using namespace std;

namespace xml
{
  bool default_value_traits<bool>::
  parse (string s, const parser& p)
  {
    if (s == "true" || s == "1" || s == "True" || s == "TRUE")
      return true;
    else if (s == "false" || s == "0" || s == "False" || s == "FALSE")
      return false;
    else
		assert(false);//throw parsing (p, "invalid bool value '" + s + "'");
	return false;
  }
}
