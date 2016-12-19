function headerRow()
{
  
  this.header = [];
  
  this.addEntry = function (header)
  {
    this.header.push(header);
  };
  
  this.getCount = function ()
  {
    return this.header.length;
  };
  
  this.ifExists = function (header)
  {
    if (this.header.indexOf(header) > -1)
    {
      return true;
    }
    else
    {
      return false;
    }
  };
  
  this.getIndex = function (header)
  {
    return this.header.indexOf(header);
  };
};


function existingreport()
{
  this.existingData = [];
  
  this.addEntry = function (name)
  {
    if (this.existingData.indexOf(name) < 0)
    {
      this.existingData.push(name);
    }
  }

  this.ifExists = function (name)
  {
    if (this.existingData.indexOf(name) > -1)
    {
      return true;
    }
    else
    {
      return false;
    }
  };
  
  this.getRow = function (name)
  {
    if (this.existingData.indexOf(name) > -1)
    {
      return this.existingData.indexOf(name)+2;
    }
    else
    {
      return -1;
    }
  };
};


function fieldServiceReport()
{
  this.item = [];
  this.value = [];
  this.valid = [];
  this.error = [];
  
  this.addEntry = function (item,value)
  {
    if (this.item.indexOf(item) < 0)
    {
      if (value == "")
      {
        this.valid.push(false);
        if (item != "Comments")
        {
          value = 0;
        }
      }
      else
      {
        this.valid.push(true);
      }
      
      this.item.push(item);
      this.value.push(value);
      this.error.push(false);
    }
  }
  
  this.setError = function(item)
  {
    if (this.item.indexOf(item) >= 0)
    {
      var idx = this.item.indexOf(item);
      this.error[idx] = true;
    }
  }

  this.ifExists = function(item)
  {
    if (this.item.indexOf(item) >= 0)
    {
      return true;
    }
    else
    {
      return false;
    }
  }
  
  this.getValue = function(item)
  {
    if (this.item.indexOf(item) >= 0)
    {
      var idx = this.item.indexOf(item);
      return this.value[idx];
    }
  }
  
  this.getCount = function()
  {
    return this.item.length;
  }
};


function validRows()
{
  this.year = [];
  this.month = [];
  this.name = [];
  this.email = [];
  this.codematch = [];
  this.row = [];
  this.addEntry = function (year,month,name,email,codematch,row)
  {
    if (this.name.indexOf(name) < 0)
    {
      this.year.push(year);
      this.month.push(month);
      this.name.push(name);
      this.email.push(email);
      this.codematch.push(codematch);
      this.row.push(row);
    }
    else
    {
      var idx = this.name.indexOf(name);
      this.row[idx] = row;
    }
  }

  this.getRow = function(name)
  {
    if (this.name.indexOf(name) >= 0)
    {
      var idx = this.name.indexOf(name);
      return this.row[idx];
    }
  }
  
  this.getCount = function()
  {
    return this.name.length;
  }
};


function userCodes()
{
  this.name = [];
  this.email = [];
  this.code = [];
  
  this.addEntry = function (name,email,code)
  {
    if (this.name.indexOf(name) < 0)
    {
      this.name.push(name);
      this.email.push(email);
      this.code.push(code);
    }
  }
  
  this.isValid = function(code)
  {
    if (this.code.indexOf(code) >= 0)
    {
      return true;
    }
    else
    {
      return false;
    }
  }
  
  this.getName = function(code)
  {
    if (this.code.indexOf(code) >= 0)
    {
      var idx = this.code.indexOf(code);
      return this.name[idx];
    }
  }

  this.getCode = function(name)
  {
    if (this.name.indexOf(name) >= 0)
    {
      var idx = this.name.indexOf(name);
      return this.code[idx];
    }
  }
  
  this.getEmail = function(name)
  {
    if (this.name.indexOf(name) >= 0)
    {
      var idx = this.name.indexOf(name);
      return this.email[idx];
    }
  }
  
  this.getCount = function()
  {
    return this.name.length;
  }
};


// Function to get user code for each email id
function getUserCodes(mailing_list)
{  
  var user_codes = new userCodes;
 
  for (var i=0;i<mailing_list.getCount();i++)
  {
    if(mailing_list.enabled[i])
    {
      var first_name = mailing_list.first_name[i];
      var last_name = mailing_list.last_name[i];
      var name = first_name + " " + last_name;
      var email = mailing_list.email[i];  
      var code = mailing_list.code[i];  
      
      user_codes.addEntry(name,email,code);
    
    }
  }
  return user_codes;
};


// Returns a random integer between min (included) and max (included)
// Using Math.round() will give you a non-uniform distribution!
function getRandomIntInclusive(min, max)
{
  return Math.floor(Math.random() * (max - min + 1)) + min;
};

function getFullMonthName(month)
{
  var fullMonthName = [];
  
  fullMonthName['Jan'] = 'January';
  fullMonthName['Feb'] = 'February';
  fullMonthName['Mar'] = 'March';
  fullMonthName['Apr'] = 'April';
  fullMonthName['May'] = 'May';
  fullMonthName['Jun'] = 'June';
  fullMonthName['Jul'] = 'July';
  fullMonthName['Aug'] = 'August';
  fullMonthName['Sep'] = 'September';
  fullMonthName['Oct'] = 'October';
  fullMonthName['Nov'] = 'November';
  fullMonthName['Dec'] = 'December';
  
  return fullMonthName[month];
};


// Get the email addresses from the Email List worksheet
function getEmailAddresses(check_column_header) 
{ 
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.setActiveSheet(spreadsheet.getSheetByName('Email List'));
  var sheet = spreadsheet.getActiveSheet();
  
  var range = sheet.getRange("A1:K200");
  var row = range.getValues();
  
  var header_row = [];
  
  for (var i=0;i<11;i++)
  {
    var header = row[0][i];
    if(header != undefined)
    {
      header_row.push(header)
    }
  }
  
  var idx_name = header_row.indexOf("NAME");
  var idx_email = header_row.indexOf("EMAIL");
  var idx_phone = header_row.indexOf("PHONE NUMBER");
  var idx_code = header_row.indexOf("USERCODE");
  var idx_check = header_row.indexOf(check_column_header);
  
  var mailing_list = new mailingList;
  
  for (var i = 1; i < row.length; i++)
  {
    var name = row[i][idx_name];
    
    if (name != undefined)
    {
      if (name != "")
      {
        var email = row[i][idx_email];
        var ph = row[i][idx_phone];
        var enabled = row[i][idx_check];
        var usercode = row[i][idx_code];
        
        var phone = ph.toString();
        var phonenumber = phone;
        
        if (phone.length == 10)
        {
          phonenumber = '(' + phone[0] + phone[1] + phone[2] + ') ' + phone[3] + phone[4] + phone[5] + '-' + phone[6] + phone[7] + phone[8] + phone[9];
        }
        
        if(enabled == "YES")
        {
          enabled = true;
        }
        else
        {
          enabled = false;
        }
        
        var names = name.split(" ");
        var first_name = names[0];
        var last_name = names[1];
        
        if(last_name == null)
        {
          last_name = "-";
        }
        
        mailing_list.addEntry(first_name,last_name,email,phonenumber,enabled,usercode);
       
      }
    }
  }
  
  return mailing_list;
  
};


// Object to store email addresses, name etc
function mailingList()
{  
  this.email = [];
  this.phone = [];
  this.first_name = [];
  this.last_name = [];
  this.name = [];
  this.enabled = [];
  this.code = [];
  
  this.addEntry = function (first_name,last_name,email,phone,enabled,code)
  {
    var name = first_name + " " + last_name;
    
    if (this.name.indexOf(name) < 0)
    {
      if(enabled)
      {
        this.email.push(email);
        this.phone.push(phone);
        this.first_name.push(first_name);
        this.last_name.push(last_name);
        this.name.push(name);
        this.enabled.push(enabled);
        this.code.push(code);
      }
    }
  };
  
  this.getCount = function ()
  {
    return this.email.length;
  };
  
  this.ifEmailExists = function (email)
  {
    if (this.email.indexOf(email) > -1)
    {
      return true;
    }
    else
    {
      return false;
    }
  };
  
  this.getIndex = function (name)
  {
    return this.name.indexOf(name);
  };
  
  this.getEmail = function(name)
  {
    if (this.name.indexOf(name) >= 0)
    {
      var idx = this.name.indexOf(name);
      return this.email[idx];
    }
  };
  
  this.getPhone = function(name)
  {
    Browser.msgBox("Name="+name);
    if (this.name.indexOf(name) >= 0)
    {
      var idx = this.name.indexOf(name);
      return this.phone[idx];
    }
  };
  
  this.isEnabled = function(name)
  {
    var index = this.name.indexOf(name);
    
    if (this.enabled[index])
    {
      return true;
    }
    else
    {
      return false;
    }
  };
};

// Function to swap first/last name
function swapName(name)
{ 
  this.names = name.split(" ");
  this.swapname = names[1] + " " + names[0];
  return swapname;
};

// Function to sort an array
function sortList(list)
{ 
  this.sorted_index = [];
  
  var slist  = list.slice(0);
  slist.sort();
  
  for (var i=0;i<slist.length;i++)
  {
    var key = slist[i];
    var idx = list.indexOf(key);
    
    this.sorted_index.push(idx);
  }
  
  return this.sorted_index;
};