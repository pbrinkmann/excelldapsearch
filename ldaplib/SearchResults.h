// Excel LDAP Search
// Copyright (C) 2008 Paul Brinkmann <paul@paulb.org>
//
// This program is free software; you can redistribute it and/or modify it
// under the terms of the GNU General Public License as published by the Free
// Software Foundation; either version 2 of the License, or (at your option)
// any later version.
//
// This program is distributed in the hope that it will be useful, but WITHOUT
// ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or 
// FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for
// more details.
//
// You should have received a copy of the GNU General Public License along with
// this program; if not, write to the Free Software Foundation, Inc., 59 Temple
//  Place, Suite 330, Boston, MA 02111-1307 USA


#pragma once

#include <vector>
#include <string>
#include <algorithm>
using std::vector;
using std::string;
using std::sort;
using std::binary_search;


//
// This makes iterating over a bunch of items really easy
//
template <typename T>
class CCollection
{
public:

	typedef auto_ptr<T> ItemPtr;

	void addItem(const T& newItem) { items.push_back(newItem); /* sort(items.begin(), items.end()); */ }
	int getItemCount() { return (int)items.size(); }
	ItemPtr getFirstItem() { collection_iterator_index=0; return getNextItem(); }
	ItemPtr getNextItem() 
	{ 
		if(items.size() == collection_iterator_index) 
			return ItemPtr(NULL); 
		
		return ItemPtr(new T(items[collection_iterator_index++])); 
	}
	ItemPtr getItemByName(const string& name)
	{ 
		// TODO: eek! eek!  this is slow - please add a hash table?  or use a sorted vector?
		// too bad binary_search just tells you if the damn thing exists and doesn't hand it back to you
		vector<T>::iterator found = find(items.begin(), items.end(), name);
		if(found != items.end())
			return ItemPtr(new T(*found));
		else
			return ItemPtr(NULL);
			
	}

private:
	vector<T> items;
	int collection_iterator_index;

};

//
// Contains the results of a search
//

class CAttribute : public CCollection<string>
{
public:

	CAttribute() : name("UNNAMED") {}
	CAttribute(const string& name_) : name(name_) {}
	string name;
	
	void addValue(const string& newValue) { addItem(newValue); }
	bool operator ==(const string& rhs) { return this->name == rhs;	}
};



class CEntry : public CCollection<CAttribute>
{
public:
	CEntry() {}
	CEntry(const string& dn_) : DN(dn_) {}
	string DN;

	void addAttribute(const CAttribute& newAttribute) { addItem(newAttribute); }
	bool operator ==(const string& rhs) { return this->DN == rhs; }
};




class CSearchResults : public CCollection<CEntry>
{
public:
	void addEntry(const CEntry& newEntry) { addItem(newEntry); }
};


typedef auto_ptr<CSearchResults> SearchResultsPtr;


