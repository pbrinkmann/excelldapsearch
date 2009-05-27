// version.hpp

//
// This code is based on the forum post at:
// 
// http://www.codeguru.com/forum/showthread.php?t=318655
//

#pragma warning(disable : 4505 4710)

#include <string>
#include <sstream>
#include <iomanip>
#include <exception>
#include <new>
#include <windows.h>


namespace version_nmsp
{
	struct language
	{
		WORD language_;
		WORD code_page_;

		language()
		{
			language_  = 0;
			code_page_ = 0;
		}
	};
}


class VersionInfo
{
public:
	VersionInfo(string appName_="")
	{
		// Get application name
		TCHAR buf[MAX_PATH] = "";

		if(appName_ != "" || ::GetModuleFileName(0, buf, sizeof(buf)))
		{
			std::string app_name;

			if(appName_ != "") {
				app_name = appName_;
			} else {
				app_name = buf;
				app_name             = app_name.substr(app_name.rfind("\\") + 1);
			}

			// Get version info
			DWORD h = 0;

			DWORD resource_size = ::GetFileVersionInfoSize(const_cast<char*>(app_name.c_str()), &h);
			if(resource_size)
			{
				resource_data_ = new unsigned char[resource_size];
				if(resource_data_)
				{
					if(::GetFileVersionInfo(const_cast<char*>(app_name.c_str()),
						0,
						resource_size,
						static_cast<LPVOID>(resource_data_)) != FALSE)
					{
						UINT size = 0;

						// Get language information
						if(::VerQueryValue(static_cast<LPVOID>(resource_data_),
							"\\VarFileInfo\\Translation",
							reinterpret_cast<LPVOID*>(&language_information_),
							&size) == FALSE)
							throw std::exception("Requested localized version information not available");
					}
					else
					{
						std::stringstream exception;
						exception << "Could not get version information (Windows error: " << ::GetLastError() << ")";
						throw std::exception(exception.str().c_str());
					}
				}
				else
					throw std::bad_alloc();
			}
			else
			{
				std::stringstream exception;
				exception << "No version information found (Windows error: " << ::GetLastError() << ")";
				throw std::exception(exception.str().c_str());
			}
		}
		else
			throw std::exception("Could not get application name");
	}

	~VersionInfo() { delete [] resource_data_; }
	std::string get_product_name() const { return get_value("ProductName"); }
	std::string get_internal_name() const { return get_value("InternalName"); }
	std::string get_product_version() const { return get_value("ProductVersion"); }
	std::string get_special_build() const { return get_value("SpecialBuild"); }
	std::string get_private_build() const { return get_value("PrivateBuild"); }
	std::string get_copyright() const { return get_value("LegalCopyright"); }
	std::string get_trademarks() const { return get_value("LegalTrademarks"); }
	std::string get_comments() const { return get_value("Comments"); }
	std::string get_company_name() const { return get_value("CompanyName"); }
	std::string get_file_version() const { return get_value("FileVersion"); }
	std::string get_file_description() const { return get_value("FileDescription"); }

private:
	unsigned char          *resource_data_;
	version_nmsp::language *language_information_;

	std::string get_value(const std::string &key) const
	{
		if(resource_data_)
		{
			UINT              size   = 0;
			std::stringstream t;
			LPVOID            value  = 0;

			// Build query string
			t << "\\StringFileInfo\\" << std::setw(4) << std::setfill('0') << std::hex
				<< language_information_->language_ << std::setw(4) << std::hex
				<< language_information_->code_page_ << "\\" << key;

			if(::VerQueryValue(static_cast<LPVOID>(resource_data_),
				const_cast<LPTSTR>(t.str().c_str()),
				static_cast<LPVOID*>(&value),
				&size) != FALSE)
				return static_cast<LPTSTR>(value);
		}

		return "";
	}
};
