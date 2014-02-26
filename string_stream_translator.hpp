#ifndef MSR_STRING_STREAM_TRANSLATOR_HPP_INCLUDED
#define MSR_STRING_STREAM_TRANSLATOR_HPP_INCLUDED

#include <string>
#include <boost/property_tree/stream_translator.hpp>

namespace msr { namespace fix {

template <class String>
class stream_translator {
    using c = typename String::value_type;
    using traits = typename String::traits_type;
    using allocator = typename String::allocator_type;
    using string = String;
    using customized = boost::property_tree::customize_stream<c, traits, String>;
    using istringstream = std::basic_istringstream<c, traits, allocator>;
    using ostreamstream = std::basic_ostringstream<c, traits, allocator>;
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
        return c('"') + oss.str() + c('"');
    }
private:
    std::locale m_loc;
};

using string_translator = stream_translator<std::string>;

using wstring_translator = stream_translator<std::wstring>;

}}

#endif
