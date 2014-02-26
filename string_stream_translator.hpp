#ifndef MSR_STRING_STREAM_TRANSLATOR_HPP_INCLUDED
#define MSR_STRING_STREAM_TRANSLATOR_HPP_INCLUDED

#include <boost/property_tree/stream_translator.hpp>

namespace msr {

template <class Char>
class boost::property_tree::stream_translator<Char, std::char_traits<Char>, std::allocator<Char>, std::basic_string<Char, std::char_traits<Char>, std::allocator<Char>>> {
    using traits = std::char_traits<Char>;
    using allocator = std::allocator<Char>;
    using string = std::basic_string<Char, traits, allocator>;
    using customized = customize_stream<Char, traits, string>;
    using istringstream = std::basic_istringstream<Char, traits, allocator>;
    using ostreamstream = std::basic_ostringstream<Char, traits, allocator>;
public:
    using internal_type = string;
    using external_type = string;
    explicit stream_translator(std::locale loc = std::locale())
        : m_loc(loc) {}
    boost::optional<external_type> get_value(const internal_type &v) {
        istringstream iss(v);
        iss.imbue(m_loc);
        string e;
        customized::extract(iss, e);
        if (iss.fail() || iss.bad() || iss.get() != traits::eof()) {
            return boost::optional<external_type>();
        }
        return e.substr(1, e.size() - 2);
    }
    boost::optional<internal_type> put_value(const string &v) {
        ostreamstream oss;
        oss.imbue(m_loc);
        customized::insert(oss, v);
        if (!oss) {
            return boost::optional<internal_type>();
        }
        return '"' + oss.str() + '"';
    }
private:
    std::locale m_loc;
};

using string_translator = boost::property_tree::stream_translator<
    char,
    std::char_traits<char>,
    std::allocator<char>,
    std::basic_string<char,
                      std::char_traits<char>,
                      std::allocator<char>>>;

}

#endif
