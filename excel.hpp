#ifndef MSR_EXCEL_HPP_INCLUDE
#define MSR_EXCEL_HPP_INCLUDE

#include <string>

namespace msr { namespace excel {

    enum class function_mask {
        unknown,
        dec2hex
    };

    function_mask get_function_mask(const std::string &str) {
        using std::string;
        string function_name;
        for (unsigned int i = 0; i < str.size(); i++) {
            if (i == '(') {
                function_name = str.substr(0, i);
            }
        }
        if (function_name == "DEC2HEX") {
            return function_mask::dec2hex;
        } else if (false) {
            // ...
        }
        return function_mask::unknown;
    }

    template <enum class function_mask>
    class function {};

    unsigned int get_row_num(const std::string &id) {
        unsigned int row_num = 0;
        for (int i = 0; i < id.size(); i++) {
            auto c = id[i];
            if (c < 'A' || c > 'Z') {
                break;
            }
            row_num *= 'Z' - 'A' + 1;
            row_num += c - 'A';
        }
        return row_num;
    }

}}

#endif
