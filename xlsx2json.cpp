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

unsigned int get_row_num(std::string id) {
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

int main(int argc, const char *argv[]) {
    using std::cout;
    using std::cerr;
    using std::endl;
    using std::string;
    using std::vector;
    using std::map;
    using std::pair;
    using std::exception;
    using std::make_pair;
    using boost::filesystem::path;
    using boost::filesystem::create_directory;
    using boost::system::error_code;
    using boost::property_tree::ptree;
    using boost::property_tree::read_xml;
    using boost::property_tree::write_json;
    using boost::lexical_cast;
    const unsigned int exit_failure = EXIT_FAILURE;
    const unsigned int exit_success = EXIT_SUCCESS;
    if (argc != 2) {
        cerr << "Usage: xlsx2json [filename]" << endl;
        return exit_failure;
    }
    path filepath = path(argv[1]);
    string filename = filepath.filename().string();
    string temp_folder = '.' + filename + ".xlsx2json";
    string output_folder = filename + "-json";
    try {
        // extract xlsx
        {
            cout << endl << "==> extracting: " << filename << endl;
            string command = "unzip " + filename + " -d " + temp_folder;
            if (system(command.c_str()) != exit_success) {
                throw exception(("Failed to extract " + filename + '.').c_str());
            }
        }
        // create directory
        {
            cout << endl << "==> creating directory..." << endl;
            error_code error;
            create_directory(output_folder, error);
            if (error) {
                throw exception("Failed to create directory.");
            }
        }
        // parse shared strings
        vector<string> strings;
        {
            cout << endl << "==> parsing: sharedStrings.xml" << endl;
            ptree root;
            read_xml(temp_folder + "/xl/sharedStrings.xml", root);
            auto &sst = root.get_child("sst");
            int i = 0;
            for (auto &si : sst) {
                if (si.first == "si") {
                    auto &e = si.second;
                    strings.push_back(e.get<string>("t"));
                    cout << i << " -> \"" << strings[i] << "\"" << endl;
                    ++i;
                }
            }
        }
        // parse workbook to get sheet informations
        map<unsigned int, string> sheets;
        {
            cout << endl << "==> parsing: workbook.xml" << endl;
            ptree root;
            read_xml(temp_folder + "/xl/workbook.xml", root);
            auto &ss = root.get_child("workbook.sheets");
            int i = 0;
            for (auto &s: ss) {
                if (s.first == "sheet") {
                    auto id = s.second.get<unsigned int>("<xmlattr>.sheetId");
                    sheets.insert({
                        id,
                        s.second.get<string>("<xmlattr>.name")
                    });
                    cout << id << " -> \"" << sheets[id] << "\"" << endl;
                    ++i;
                }
            }
        }
        // parse each sheet
        for (auto &sheet: sheets) {
            cout << endl << "==> parsing: sheet" << sheet.first << ".xml" << endl;
            ptree xml_root;
            vector<string> keys;
            read_xml(temp_folder + "/xl/worksheets/sheet" +
                     lexical_cast<string>(sheet.first) + ".xml",
                     xml_root);
            // get titles
            auto &data = xml_root.get_child("worksheet.sheetData");
            {
                cout << "[title]" << endl;
                auto &first_row = begin(data)->second;
                int i = 0;
                for (auto &c: first_row) {
                    if (c.first == "c") {
                        auto num = c.second.get<unsigned int>("v");
                        keys.push_back(strings[num]);
                        cout << i << " -> \"" << keys[i] << "\"" << endl;
                        ++i;
                    }
                }
            }
            ptree json_root;
            // get contents
            {
                cout << "[contents]" << endl;
                for (auto &r: data) {
                    if (r.first == "row" &&
                        r.second.get<unsigned int>("<xmlattr>.r") != 1) {
                        ptree dictionary;
                        for (auto &c: r.second) {
                            if (c.first == "c") {
                                auto id = c.second.get<string>("<xmlattr>.r");
                                auto row_num = get_row_num(id);
                                try {
                                    auto t = c.second.get<string>("<xmlattr>.t");
                                    if (t == "s") {
                                        auto v = c.second.get<unsigned int>("v");
                                        dictionary.put(keys[row_num], strings[v], msr::string_translator());
                                        cout << keys[row_num] << " -> \""
                                             << strings[v] << "\"" << endl;
                                    }
                                } catch (exception &) {
                                    auto v = c.second.get<long double>("v");
                                    dictionary.put(keys[row_num], v);
                                    cout << keys[row_num] << " -> " << v << endl;
                                }
                            }
                        }
                        json_root.push_back({"", dictionary});
                    }
                }
            }
            // for solving boost::property_tree::json_parser bug
            // ptree new_root;
            // new_root.put_child("root", json_root);
            // output
            {
                cout << endl << "==> converting sheet: " << sheets[sheet.first] << endl;
                write_json(output_folder + "/" + sheets[sheet.first] + ".json", json_root);
            }
        }
        cout << endl << "==> finished." << endl << endl;
    } catch (exception &e) {
        cerr << e.what() << endl << endl;
        return exit_failure;
    }
    system(("rm -rf " + temp_folder).c_str());
    return exit_success;
}
