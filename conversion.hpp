#ifndef MSR_CONVERSION_HPP_INCLUDE
#define MSR_CONVERSION_HPP_INCLUDE

#define _CRT_SECURE_NO_WARNINGS
#include <vector>
#include <string>
#include <cstdlib>
#include <locale>
#include <codecvt>

namespace msr {
    namespace conversion {

        std::string to_string(const std::wstring &wstr) {
            auto size = wstr.size() * 2 + 1;
            char *_str;
            _str = new char[size];
            // setlocale(LC_ALL, "");
            wcstombs(_str, wstr.c_str(), size);
            std::string str = _str;
            delete _str;
            return str;
        }

        std::wstring to_wstring(const std::string &str) {
            auto size = str.size() + 1;
            wchar_t *_wstr;
            _wstr = new wchar_t[size];
            // setlocale(LC_ALL, "");
            mbstowcs(_wstr, str.c_str(), size);
            std::wstring wstr = _wstr;
            delete _wstr;
            return wstr;
            //std::wstring_convert<std::codecvt_utf8<char>, char> utf8;
            //std::wstring wstr = utf8.from_bytes(str);
        }
    }
}
#endif
