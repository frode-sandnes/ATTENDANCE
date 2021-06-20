// By Frode Eika Sandnes, Oslo Metropolitan University, Oslo, Norway, May 2021

// Global variables
var threshold = 0.5; // leanient
//	var threshold = 0.8; // strict
var thresholdLecture = 5;
var dates;
var names;
var keywords;
var days = [];
var times = [];
var url = "";

// called when page is loaded
function onLoad()
	{
	if (typeof console.log(document.getElementById("URLID").value) !== 'undefined')
		{
		url = console.log(document.getElementById("URLID").value);
		}
	else if (window.location.search.length > 0)
		{
		var parts = window.location.search.split("=");
		url = parts[1];
		}
	
	// get sheet		
	if (url.length > 0)
		{
		// cut away the heading characters
		try
			{
			var parts = url.split("%2Fd%2F"); // "/d/"
			url = parts[1];
			// cut away the trailing characters
			parts = url.split("%2F");
			url = parts[0];
			}
		catch (err)
			{
			document.getElementById("output1Id").innerHTML = "<p>Invalid URL.</p>" 
			return false;
			}
			
		// try to read thie file
		try
			{
			getSheet();
			}
		catch (err)
			{
			document.getElementById("output1Id").innerHTML = "<p>Problem with the Google Sheets file - try sharing it.</p>" 
			}
		}
	else
		{
		// error mesage
		document.getElementById("output1Id").innerHTML = "<p>A reference to valid and publically shared Google Sheets file is needed.</p>";	
		}
		
	return false;
	}
// download from Google sheets
function getSheet()
	{
    new RGraph.Sheets(url, 
		function (sheet)
			{
			// get data and put into globals
			dates = sheet.get('A');
			names = sheet.get('B');
			keywords = sheet.get('C');

			// remove the headers - i.e. remove the first element in each array
			dates.splice(0, 1);			
			names.splice(0, 1);			
			keywords.splice(0, 1);			
			
			// split dates into days and times
			for (var s of dates)
				{
				var strs = s.split(" kl. ");  //  may have to replace this for different locale
				days.push(strs[0]);
				times.push(strs[1]);
				}
	
			// when file is loaded - process the data (could take some time)
			process();
			});
	}
// string distance measure using DICE
function dice(l, r) 
	{
	if (l.length < 2 || r.length < 2) return 0;
	let lBigrams = new Map();
	for (let i = 0; i < l.length - 1; i++) 
		{
		const bigram = l.substr(i, 2);
		const count = lBigrams.has(bigram)
		? lBigrams.get(bigram) + 1
		: 1;
		lBigrams.set(bigram, count);
		};
	let intersectionSize = 0;
	for (let i = 0; i < r.length - 1; i++) 
		{
		const bigram = r.substr(i, 2);
		const count = lBigrams.has(bigram)
		? lBigrams.get(bigram)
		: 0;
		if (count > 0) 
			{
			lBigrams.set(bigram, count - 1);
			intersectionSize++;
			}
		}
	return (2.0 * intersectionSize) / (l.length + r.length - 2);
	}	
// finds the unique set of strings in list based on approximate string matching
function findUniqueSet(list)
	{
	var unique = [];
	for (var s of list)
		{
		s = s.toLowerCase();
		// search approximatesly in existing list
		var exists = false;
		for (var t of unique)
			{
			if (dice(s,t) > threshold)
				{
				exists = true;	
				break				
				}
			}
		if (!exists)	// must be inserted once, and must have enough number of occurrences
			{
			var f = approximateFrequency(s,names);
			if (f > thresholdLecture)
				{
				unique.push(s);
				}
			}
		}
	return unique;
	}
// return the number of times an item occur in the lsit, changes to lowercase and also remove spaces	
function frequency(item,list)
	{
	var count = 0;
	for (var s of list)
		{
		s = s.toLowerCase();			
		s = s.split(' ').join('');		
		if (s.match(item))
			{
			count++;
			}			
		}
	return count;
	}
// return the number of times an item occur in the lsit, changes to lowercase and also remove spaces	
function approximateFrequency(item,list)
	{
	var count = 0;
	for (var s of list)
		{
		s = s.toLowerCase();			
		if (dice(s,item) > threshold)			
			{
			count++;
			}
		}
	return count;
	}	
// return the frequency of each untique item	
function histogram(list)
	{
	var m = new Map();		// find unique items
	var s = new Set(list);
	var a = Array.from(s);
	for (var t of a)
		{	
		t = t.toLowerCase();
		t = t.split(' ').join('');
		if (t === "")
			{
			continue;
			}
		var f = frequency(t,list);
		if (f > thresholdLecture)
			{
			m.set(t,f);
			}
		}
	return m;
	}
// return the frequency of each untique item with minimum number of entries	
function filter(list)
	{
	var m = [];		// find unique items
	var s = new Set(list);
	var a = Array.from(s);
	for (var t of a)
		{	
		t = t.toLowerCase();
		t = t.split(' ').join('');		
		var f = frequency(t,list);
		if (f > thresholdLecture)
			{
			m.push(t);
			}
		}
	return m;
	}	
// find the starttimes based on first occurrance of valid date was recorded	
function findStartTimes(uniqueDays)
	{
	var startTimes = [];
	let daylist = Array.from(uniqueDays.keys());	
	for (var ud of daylist)
		{
		var index = days.indexOf(ud);
		startTimes.push(times[index]);
		}
	return startTimes;
	}	
function process()
	{
	var uniqueNames = findUniqueSet(names);
	var uniqueDays = histogram(days);
	var uniqueCodes = histogram(keywords);
	var startTimes = findStartTimes(uniqueDays);
	buildStructure(uniqueNames,uniqueDays,uniqueCodes,startTimes);
	}
// search a list to see if any is sufficiently similar	
function approxInclude(item,list)
	{
	for (var s of list)
		{
		if (dice(s,item) > threshold)
			{
			return true;	// we found a match above limit
			}
		}
	return false;	// no match above limit
	}
// find the indes of the most appropriate map	
function approxInindexOf(item,list)
	{
	for (var s of list)
		{
		if (dice(s,item) > threshold)
			{
			return list.indexOf(s);	// we found a match above limit
			}
		}
	return -1;	// no match above limit
	}	
function timeStringToSecs(t)
	{
	// split into parts
	var ta = t.split(".");
	// return the value by adding hours, minutes and seconds
	return 3600*+ta[0]+60*ta[1]+1*ta[2];
	}
function timeDifference(t1,t2)
	{
	var d1 = timeStringToSecs(t1);		
	var d2 = timeStringToSecs(t2);
	return d2-d1;
	}
function buildStructure(uniqueNames,uniqueDays,uniqueCodes,startTimes)
	{
	// get array versjon of the frequency maps
	var ud = Array.from(uniqueDays.keys());		
	var codes = Array.from(uniqueCodes.keys());			

	// declare a two-dimensional structure for the result
	var record = new Array(uniqueNames.length);
	for (var i = 0; i < record.length; i++) 
		{
		record[i] = new Array(ud.length);
		}
		
	// traverse the full list find valid lectures
	for (var i = 0;i<names.length;i++)
		{
		var n = names[i];
		var c = keywords[i];
		var d = days[i];
		var t = times[i];
		// aborting cases - invalid day
		
		if (!ud.includes(d))
			{
			continue;
			}
		// find the lecture index since it is a valid days
		var idx = ud.indexOf(d);	
		// aborting cases - invalid keyword
		if (!approxInclude(c,codes))
			{
			continue;
			}
			
		// aborting cases - invalid person
		if (!approxInclude(n,uniqueNames))
			{
			continue;
			}				

		// uncomment if the keyword to be identical - many make mistakes
/*		if (dice(c,codes[idx]) < threshold)
			{
			continue;
			}*/
		// is the registration within time?
		var st = startTimes[idx];
		var delay = timeDifference(st,t);
		var idx1 = approxInindexOf(n,uniqueNames);
		record[idx1][idx] = delay;
		}

		
	// create output of the results
	// attendance statistics
	/////////////////////////////////////////
	
	var out = "";		
	    out += "<h2>Attendance statistics per student</h2>";	
		out += "<table>";
		// headers
		out += "<tr>";	
			out += "<th>";
				out += "Name"
			out += "</th>";	
			
			out += "<th>";
				out += "attendance count";
			out += "</th>";	
			
			out += "<th>";
				out += "Persentage %";
			out += "</th>";	
			
		out += "</tr>";		
	// create attendance count
	for (var row of record)
		{
		out += "<tr>";			
		// the name in first column		
			out += "<td>";		
				out += uniqueNames[record.indexOf(row)];
			out += "</td>";		

			out += "<td>";

		var count = 0;	
		for (var cell of row)
			{
			if (typeof cell !== 'undefined')	
				{
				count++;
				}						
			}
			out += count+"/"+row.length;
			out += "</td>";	

			out += "<td>";	
			out += (100*count/row.length).toFixed(1);
			out += "</td>";				
		out += "<tr>";					
		}	
	out += "</table>";	

	
	// class attendance statistics
	/////////////////////////////////////////
	
	    out += "<h2>Attendance statistics per class</h2>";	
		out += "<table>";
		// headers
		out += "<tr>";	
			out += "<th>";
				out += "Date"
			out += "</th>";	
			
			out += "<th>";
				out += "attendance count";
			out += "</th>";	

			out += "<th>";
				out += "Persentage %";
			out += "</th>";				
		out += "</tr>";		
	// create attendance count
	for (var i=0;i<record[0].length;i++)
		{
		out += "<tr>";			
		// the name in first column		
			out += "<td>";		
				out += ud[i];
			out += "</td>";		

			out += "<td>";

		var count = 0;	
		for (var j=0;j<record.length;j++)
			{
			var cell = record[j][i];
			if (typeof cell !== 'undefined')	
				{
				count++;
				}						
			}
			out += count;
			out += "</td>";	
			out += "<td>";	
			out += (100*count/uniqueNames.length).toFixed(1);
			out += "</td>";	
		out += "<tr>";					
		}	
	out += "</table>";		
	

	// Attendance details
	/////////////////////////////////////////
	
	    out += "<h2>Attendance details (delay in mins.)</h2>";	

		out += "<table>";
	// first headers
		out += "<tr>";
			// blank first ityem
			out += "<th>";
			out += "</th>";
		for (var q of ud)
			{
			out += "<th>";
				out += q.slice(0, -5);
			out += "</th>";
			}
		out += "</tr>";
		
		// then the body
		// the delays
		for (var row of record)
			{
			out + "</tr>";				
			// the name in first column
				out += "<td>";		
					out += uniqueNames[record.indexOf(row)];
				out += "</td>";		

			for (var cell of row)
				{
				out += "<td>";
				if (typeof cell !== 'undefined')	
					{
					out += (cell/60).toFixed(1);	// show in minutes
					}						
				
				out += "</td>";				
				}
			out += "<tr>";					
			}		
	out += "</tr>";
	out += "</table>";
	
	
	// output keywords:
	/////////////////////////////////////////

    out += "<p/>";	
    out += "<h2>Class details</h2>";	
		
	out += "<table>";	

		// headers
		out += "<tr>";	
			out += "<th>";
				out += "Date"
			out += "</th>";	
			
			out += "<th>";
				out += "Keyword"
			out += "</th>";	

			out += "<th>";
				out += "Time of first response"
			out += "</th>";	

		out += "</tr>";

		for (var i=0;i<ud.length;i++)
			{
		out += "<tr>";
				out += "<td>";
				out += ud[i];		
				out += "</td>";			

				out += "<td>";
				out += codes[i];			
				out += "</td>";			

				out += "<td>";
				out += startTimes[i];			
				out += "</td>";			
				
			out += "</tr>";
			}
	out += "</table>";
	

	// timing analysis
	var meanTimes = timeAnalysis(uniqueNames,record);
	out += "<h2>Time analysis (median)</h2>";
	out = dumpMapHTML(out,meanTimes);
	
	
	// Add error logs
	/////////////////////////////////////////
	var repeatedNames = uniqueVariations(uniqueNames,names);
	out += "<h2>Name variations</h2>";
	out = dumpMapArrayHTML(out,repeatedNames);

	var repeatedCodes = uniqueVariations(codes,keywords);
	out += "<h2>Code variations</h2>";
	out = dumpMapArrayHTML(out,repeatedCodes);
	
	var illegalDates = dateDeviations(ud,days,names);
	out += "<h2>Illegal dates</h2>";
	out = dumpMapArrayHTML(out,illegalDates);	
	
	var error = errorSwapInput(codes,uniqueNames,keywords,names,days)
	out += "<h2>Erronously swapped input</h2>";
	out = dumpMapHTML(out,error);	
	
	// output table check info
	var meanCodes = listSpread(ud, days, codes, keywords, threshold);	
	var meanNames = listSpread(ud, days, uniqueNames, names, threshold);	
	out += "<h2>Table alignment check</h2>"
	out += "<p>Mean "+meanCodes.toFixed(1)+" keywords per class (should be close to 1)</p>";
	out += "<p>mean "+meanNames.toFixed(1)+" names per class (should be close to the class size)</p>";

	// put result in the html document
	document.getElementById("output1Id").innerHTML = out;	
	}

// find list of unique variations
function uniqueVariations(ulist,flist)
	{
	var result = new Map();
	for (var s1 of ulist)
		{
		s1 = s1.toLowerCase();
		var variations = [];
		for (var s2 of flist)
			{
			s2 = s2.toLowerCase();
			if (s1.match(s2))
				{
				continue;
				}
			if (s1.match(s2.split(' ').join('')))
				{
				continue;
				}						
			if (dice(s1,s2) > threshold)
				{
				if (!variations.includes(s2))
					{
					variations.push(s2);
					}
				}
			}
		if (variations.length > 0)
			{
			result.set(s1, variations);
			}
		}		
	return result;
	}
// output helper	
function dumpMapArrayHTML(out,map)
	{
	for (var key of map.keys())
		{
		out += "<p>";
		out += "<b>"+key+":</b>"
		for (var val of map.get(key))
			{
			out += "\""+val +"\", ";
			}
		out += "</p>";
		}
	return out;
	}
function dumpMapHTML(out,map)
	{
	for (var key of map.keys())
		{
		out += "<p>";
		out += "<b>"+key+":</b>"+map.get(key);
		out += "</p>";
		}
	return out;
	}	
// list students who submitted entry on a non-class day -warning	
function dateDeviations(ud,days,names)
	{
	var result = new Map();		
	for (var d of days)
		{			
		if (!ud.includes(d))
			{
			var list = [];
			if (result.has(d))
				{
				list = result.get(d);
				}
			var idx = days.indexOf(d);
			list.push(names[idx]);
			result.set(d,list);
			}
		}
	return result;
	}
// find entries where the keywords and names have been swapped	
function errorSwapInput(codes,uniqueNames,keywords,names,days)
	{
	var result = new Map();			
	for (var i=0; i< names.length;i++)
		{
		var name = names[i];
		var code = keywords[i];
		if (approxInclude(name,codes) && approxInclude(code,uniqueNames))
			{
			result.set("Name:"+name+", code:"+code,days[i]);
			}
		}
	return result;		
	}	
// output spreadsheet
function outputSpreadsheet()
	{
	if (typeof linked === 'undefined')	
		{
		return false;
		}		
    var filename='linked-output.xlsx';
	var ws = XLSX.utils.json_to_sheet(linked);
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Linked-output");
    XLSX.writeFile(wb,filename);
	return false;
	}	
// time analysis (medians as robust to outliiers)
function timeAnalysis(uniqueNames,record)
	{
	var result = new Map();
	for (var row of record)
		{
		var list = [];
		for (var cell of row)
			{
			if (typeof cell !== 'undefined')	
				{
				list.push(cell);			
				}	
			}
		var median = -1;
		if (list.length>0)
			{
			list.sort(function(a, b){return a - b});	// sort in order	
			var midPoint = Math.floor(list.length/2);
			median = list[midPoint];	// middle element
			}	
		result.set(uniqueNames[record.indexOf(row)],(median/60).toFixed(1));			
		}		
	return result;
	}		
// checking if form fields are in the right order
// check spread of list entries per day
function listSpread(ud, days, referenceList, list, thrshld)
	{
	// first establish list of unique items per valid day
	var m = new Map();
	for (var item of list)
		{
		var idx = list.indexOf(item);
		var day = days[idx];
		if (ud.includes(day))
			{
			var idx2 = ud.indexOf(day);
			var referenceItem = referenceList[idx2];
			if (dice(item,referenceItem) < thrshld)
				{
				if (!m.has(day))	// check if it is already there
					{
					m.set(day,0);
					}
				var v = m.get(day);	// increment one
				v++;
				m.set(day,v);
				}
			}
		}
	// then go through map and calculate average
	var mean = 0;
	var count = 0;
	for (var v of m.keys())
		{
		mean += m.get(v);	
		count++;
		}		
	mean /= count;
	return mean;
	}


