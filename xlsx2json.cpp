
#include "conversion.hpp"
#include "excel.hpp"
#include <iostream>
#include <string>
#include <vector>
#include <map>
#include <boost/filesystem.hpp>
#include <boost/property_tree/ptree.hpp>
#include <boost/property_tree/xml_parser.hpp>
#include "json_parser_write.hpp"
#include "string_stream_translator.hpp"
#include <boost/property_tree/json_parser.hpp>
#include <boost/lexical_cast.hpp>
#include <boost/locale.hpp>




int main(int argc, const char *argv[]) {
    using std::string;
    using std::wstring;
    using std::vector;
    using std::map;
    using std::pair;
    using std::exception;
    using std::locale;
    using std::cout;
    using std::wcout;
    using std::wcerr;
    using std::endl;
    using std::make_pair;
    using boost::filesystem::wpath;
    using boost::locale::generator;
    using boost::system::error_code;
    using boost::property_tree::wptree;
    using boost::filesystem::create_directory;
    using boost::property_tree::read_xml;
    using boost::property_tree::write_json;
    using boost::lexical_cast;
    using msr::fix::string_translator;
    using msr::fix::wstring_translator;
    using msr::conversion::to_string;
    using msr::conversion::to_wstring;
    using msr::excel::get_row_num;
    using msr::excel::get_function_mask;
    using msr::excel::function_mask;
    const unsigned int exit_failure = EXIT_FAILURE;
    const unsigned int exit_success = EXIT_SUCCESS;
    generator generate;
    // auto system_locale = generate("");
    // auto origin = locale::global(system_locale);
    wcout.imbue(generate(""));
    std::ios_base::sync_with_stdio(false);
    if (argc != 2) {
        wcerr << L"Usage: xlsx2json [filename]" << endl;
        return exit_failure;
    }
    string argv_1 = argv[1];
    wpath filepath = wpath(to_wstring(argv_1));
    wstring filename = filepath.filename().wstring();
    wstring temp_folder = L".xlsx2json";
    wstring output_folder = filename + L"-json";
    try {
        // extract xlsx
        {
            wcout << endl << L"==> extracting: " << filename << endl;
            wstring command = L"unzip -qq -o " + filename + L" -d " + temp_folder;
            if (system(to_string(command).c_str()) != exit_success) {
                throw exception(to_string(L"Failed to extract " + filename + L'.').c_str());
            }
        }
        // create directory
        {
            wcout << endl << L"==> creating directory..." << endl;
            error_code error;
            create_directory(to_string(output_folder), error);
            if (error) {
                throw exception("Failed to create directory.");
            }
        }
        // parse shared strings
        vector<wstring> strings;
        {
            wcout << endl << L"==> parsing: sharedStrings.xml" << endl;
            wptree root;
            read_xml(to_string(temp_folder + L"/xl/sharedStrings.xml"), root);
            auto &sst = root.get_child(L"sst");
            int i = 0;
            for (auto &si : sst) {
                if (si.first == L"si") {
                    auto &e = si.second;
                    strings.push_back(e.get<wstring>(L"t"));
                    wcout << i << L" -> \"" << strings[i] << L"\"" << endl;
                    ++i;
                }
            }
        }
        // parse workbook to get sheet informations
        map<unsigned int, wstring> sheets;
        {
            wcout << endl << L"==> parsing: workbook.xml" << endl;
            wptree root;
            read_xml(to_string(temp_folder + L"/xl/workbook.xml"), root);
            auto &ss = root.get_child(L"workbook.sheets");
            int i = 0;
            for (auto &s: ss) {
                if (s.first == L"sheet") {
                    auto id = s.second.get<unsigned int>(L"<xmlattr>.sheetId");
                    sheets.insert({
                        id,
                        s.second.get<wstring>(L"<xmlattr>.name")
                    });
                    wcout << id << L" -> \"" << sheets[id] << L"\"" << endl;
                    ++i;
                }
            }
        }
        // parse each sheet
        for (auto &sheet: sheets) {
            wcout << endl << L"==> parsing: sheet" << sheet.first << L".xml" << endl;
            wptree xml_root;
            vector<wstring> keys;
            read_xml(to_string(temp_folder + L"/xl/worksheets/sheet" +
                     lexical_cast<wstring>(sheet.first) + L".xml"),
                     xml_root);
            // get titles
            auto &data = xml_root.get_child(L"worksheet.sheetData");
            {
                wcout << L"[title]" << endl;
                auto &first_row = begin(data)->second;
                int i = 0;
                for (auto &c: first_row) {
                    if (c.first == L"c") {
                        auto num = c.second.get<unsigned int>(L"v");
                        keys.push_back(strings[num]);
                        wcout << i << L" -> \"" << keys[i] << L"\"" << endl;
                        ++i;
                    }
                }
            }
            wptree json_root;
            // get contents
            {
                wcout << "[contents]" << endl;
                for (auto &r: data) {
                    if (r.first == L"row" &&
                        r.second.get<unsigned int>(L"<xmlattr>.r") != 1) {
                        wptree dictionary;
                        for (auto &c: r.second) {
                            if (c.first == L"c") {
                                auto id = c.second.get<wstring>(L"<xmlattr>.r");
                                auto row_num = get_row_num(to_string(id));
                                try {
                                    auto t = c.second.get<wstring>(L"<xmlattr>.t");
                                    if (t == L"s") {
                                        auto v = c.second.get<unsigned int>(L"v");
                                        dictionary.put(keys[row_num], strings[v], wstring_translator());
                                        wcout << keys[row_num] << L" -> \""
                                              << strings[v] << L"\"" << endl;
                                    } else if (t == L"str") {
                                        auto f = c.second.get<wstring>(L"f");
                                        if (get_function_mask(to_string(f)) == function_mask::dec2hex) {
                                            auto v = c.second.get<long double>(L"v");
                                            dictionary.put(keys[row_num], v);
                                            wcout << keys[row_num] << L" -> "
                                                  << v << endl;
                                        } else if (false) {
                                            // ...
                                        }
                                    }
                                } catch (exception &) {
                                    try {
                                        auto v = c.second.get<long double>(L"v");
                                        dictionary.put(keys[row_num], v);
                                        wcout << keys[row_num] << L" -> " << v << endl;
                                    } catch (exception &e) {
                                        wcout << to_wstring(e.what()) << endl;
                                    }
                                }
                            }
                        }
                        if (!dictionary.empty()) {
                            json_root.push_back({L"", dictionary});
                        }
                    }
                }
            }
            // for solving boost::property_tree::json_parser bug
            // ptree new_root;
            // new_root.put_child("root", json_root);
            // output
            {
                wcout << endl << L"==> converting sheet: " << sheets[sheet.first] << endl;
                write_json(to_string(output_folder + L"/" + sheets[sheet.first] + L".json"), json_root);
            }
        }
        wcout << endl << L"==> finished." << endl << endl;
    } catch (exception &e) {
        string what = e.what();
        wcerr << to_wstring(e.what()) << endl << endl;
        return exit_failure;
    }
    // system(to_string(L"rm -rf " + temp_folder).c_str());
    return exit_success;
}
