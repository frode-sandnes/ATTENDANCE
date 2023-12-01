// By Frode Eika Sandnes, Oslo Metropolitan University, Oslo, Norway, May 2021
// revised Dec, 2022.
// revised Oct, 2023.

"use strict";

// Global variables
const threshold = 0.5; // lenient
const lateRegistrationDelay = 30; // minutes
let absenceCode = "";
let minAttendanceLimit = 0;
let XL_row_object = "";
let lastButtonId = "buttonstudentOverview";
let lastId = "studentOverview";

function loadSpreadSheet(event)
	{
	const files = event.target.files;

	for (let i = 0, f; f = files[i]; i++) 
		{			
		let reader = new FileReader();
		
		reader.onload = (function(theFile) 
			{
			return function(e) 
				{
				const workbook = XLSX.read(e.target.result, {type: 'binary'});	
				for (let sheetName of workbook.SheetNames)
					{
					XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
					processSheet(XL_row_object);
					}						
				};
			})(f);		
		reader.readAsBinaryString(f);
		}
	}

// globals
let tables = new Map();
let titles = new Map();
// utility to get excel dump
function saveExcelSheet()
    {
	var ws = XLSX.utils.json_to_sheet(tables.get(lastId));
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Extracted-grades");
    XLSX.writeFile(wb,titles.get(lastId) + ".xlsx");
    }	

function prefix()
	{
	const {pathname} = window.location;
	return pathname;
	}	
function privateLocalStoreGetItem(key)
	{
	return localStorage.getItem(prefix() + key); 
	}
function privateLocalStoreRemoveItem(key)
	{
	localStorage.removeItem(prefix() + key);	
	}	
function privateLocalStoreSetItem(key, value)
	{
	localStorage.setItem(prefix() + key, value);	
	}

function processSheet(XL_row_object)
    {
	if (XL_row_object == null)
		{
		alert("Problem loading sheet");
//		wipeData();
		return;
		}
	// Assign columns
	let allColumnHeaders = [...Object.keys(XL_row_object[0])];

//	wipeData();			
	// insert html for form assignments
	createHtmlList("timestampInput", allColumnHeaders);
	createHtmlList("fullnameInput", allColumnHeaders);
	createHtmlList("codeInput", allColumnHeaders);
	populateForm(allColumnHeaders);
	if (!validColumnsSelected())
		{
		return;	// abort
		}	
	// populate the collumnHeaders object
	let columnHeaders = [document.getElementById("timestampColumn").value, 
						 document.getElementById("nameColumn").value, 
						 document.getElementById("codeColumn").value ];
	// find the date format - based on first regitation
	// could be problem with csv file and parsing of data formats, therefore export to excel first
	const datetimeFormat = extractTimestamp(XL_row_object[0][columnHeaders[0]]);					
	// built key structure	
	let registrations = XL_row_object.map(e => ({datestamp:e[columnHeaders[0]], 
												 name:((e[columnHeaders[1]] !== undefined)?e[columnHeaders[1]].toLowerCase():""), 
												 keyword:((e[columnHeaders[2]] !== undefined)?e[columnHeaders[2]].toLowerCase():""),
												 day:parseDateTime(e[columnHeaders[0]],datetimeFormat).date,
												 time:getTimeString(e[columnHeaders[0]], datetimeFormat)
												 }));										 
	keepData();	// keep the data in browser storage for next time	
	buildStructure(registrations);
	}
function validColumnsSelected()
	{
	return document.getElementById("timestampColumn").value.length > 0 &&
		   document.getElementById("nameColumn").value.length > 0 &&
		   document.getElementById("codeColumn").value.length > 0;
	}
function populateParameter(form, storage, allColumnHeaders)
	{
	let storedValue = privateLocalStoreGetItem(storage); 	
	if (allColumnHeaders.includes(storedValue))
		{
		document.getElementById(form).value = storedValue;
		}
	}
//	helperClearForm();			// uncomment during debug
function helperClearForm()
	{
	privateLocalStoreRemoveItem("timestampColumn");	
	privateLocalStoreRemoveItem("nameColumn");	
	privateLocalStoreRemoveItem("codeColumn");	
	}
function helperClearFields()
	{
	document.getElementById("timestampColumn").value = "";	
	document.getElementById("nameColumn").value = "";	
	document.getElementById("codeColumn").value = "";	
	}
function populateForm(allColumnHeaders)
	{
	helperClearFields();		
	populateParameter("timestampColumn", "timestampColumn",allColumnHeaders);
	populateParameter("nameColumn", "nameColumn",allColumnHeaders);
	populateParameter("codeColumn", "codeColumn",allColumnHeaders);		
	}	
function updateParameter(form, storage)
	{
	let formValue = document.getElementById(form).value;
	let storedValue = privateLocalStoreGetItem(storage); 	
	if (formValue !== storedValue)
		{
		privateLocalStoreSetItem(storage, formValue);			
		}
	}
function updateList()
	{
	updateParameter("timestampColumn", "timestampColumn");
	updateParameter("nameColumn", "nameColumn");
	updateParameter("codeColumn", "codeColumn");
	if (validColumnsSelected())
		{
		processSheet(XL_row_object);
		}
	}
function createHtmlList(id, headers)
	{
	// clean old dropDown menu
	document.getElementById(id)?.remove();	// remove existing drop down with optional chaining
	let datalist  = document.createElement('datalist');
	datalist.id = id;
	document.body.appendChild(datalist);
	headers.forEach(header => 
		{
		// insert name into the datalist
		let option = document.createElement('option');  
		option.value = header;
		datalist.appendChild(option);  
		});
	}
function getTimeString(s, datetimeFormat)
	{
	const d = parseDateTime(s, datetimeFormat);
	return zeroPad(d.hour,2)+"."+zeroPad(d.minute,2)+".00";
	}
function extractTimestamp(dateString)
	{
	const digitGroups = dateString.match(/\d{1,4}/g); 
	let separators = dateString.match(/\D+/g);   // // et all the separating characters

	// ensure that there is always a prefix character - insert dummy separator if digits start before separator
	if (dateString.indexOf(separators[0])>dateString.lastIndexOf(digitGroups[0]))
		{
		separators = ["",...separators];
		}

	let classification = [];
	// deduce according to separator contents
	separators.forEach(separator => 
		{
		if (digitGroups.length == classification.length)
			{
			return;  // no need to look further
			}
		if (separator.includes("/"))
			{
			if (classification.length>0)
				{
				classification.pop();     // backtrack and redo               
				classification = [...classification, "date"];    
				}
			classification = [...classification, "date"];    
			}
		else if (separator.includes(":"))
			{
			if (classification.length>0)
				{
				classification.pop();       // backtrack and redo                    
				classification = [...classification, "time"];    
				}                    
			classification = [...classification, "time"];    
			}
		else
			{
			classification = [...classification, "unknown"];                        
			}
		});
	// fill in the rest if one is missing
	if (classification.indexOf("date") == -1)
		{
		classification = classification.map(e => (e === "unknown")?"date" : e);
		}
	else if (classification.indexOf("time") == -1)
		{
		classification = classification.map(e => (e === "unknown")?"time" : e);
		}       

	let idx = classification.indexOf("time");
	classification[idx] = "hour";
	if (classification[idx+1].match("time"))
		{
		classification[idx+1] = "minute";
		}
	if (classification.length > idx+2 && classification[idx+2].match("time"))
		{
		classification[idx+2] = "second";
		}
	// detect year if four digit code
	digitGroups.forEach((e, i) => {if (e.length > 3) classification[i] = "year"});
	return classification;
	}
function parseDateTime(dateString, classification)
	{
	let date = "",hour = 0, minute = 0;
	const digitGroups = dateString.match(/\d{1,4}/g); 
	classification.forEach((c, i) => 
		{
		if (c.match("date"))
			{
			if (date.length > 0)
				{
				date += "/";        // add delimiter
				}
			date += digitGroups[i];
			}
		if (c.match("hour"))
			{
			hour = digitGroups[i];
			}
		if (c.match("minute"))
			{
			minute = digitGroups[i];
			}
		});
	return {date:date, hour:hour, minute: minute};    
	}
const zeroPad = (num, places) => String(num).padStart(places, '0');
// use unigrams if strings are short
function unigramDice(str1,str2)
	{
	// character by character compare
	const matches = str1.split('').reduce((accumulator, value, index) => 
		{
		if (index < str2.length)
			{
			if (value == str2.charAt(index))
				{
				accumulator++;
				}	
			}
		return accumulator;
		}, 0);
	const score = matches / Math.max(str1.length, str2.length);
	return score;
	}
// string distance measure using DICE
function dice(str1, str2)
	{
	if (typeof str1 == 'undefined' || typeof str2 == 'undefined' || str1.length < 2 || str2.length < 2) return 0;
	if (Math.max(str1.length, str2.length) < 5)
		{
		return unigramDice(str1, str2);
		}
	const charArr1 = [...str1.toLowerCase()], charArr2 = [...str2.toLowerCase()];
	const bigrams1 = charArr1.filter((e, i) => i < charArr1.length - 1)
						 .map((e, i) => e + charArr1[i + 1]);						
	const bigrams2 = charArr2.filter((e, i) => i < charArr2.length - 1)
						 .map((e, i) => e + charArr2[i + 1]);
	const intersection = new Set(bigrams1.filter(e => bigrams2.includes(e)));
	// count number of intersecting bigrams
	const intersectionCounts = [...intersection].map(bigram => Math.min(bigrams1.filter(e => e == bigram).length,
							 								            bigrams2.filter(e => e == bigram).length));
	const intersectionSize = intersectionCounts.reduce((accumulator, e) => accumulator + e, 0);
	return (2.0 * intersectionSize) / (str1.length + str2.length - 2);
	}
// string distance measure using DICE
function asymmetricDice(str1, str2)
	{
	if (str1.length < 2 || str2.length < 2) return 0;
	const charArr1 = [...str1.toLowerCase()], charArr2 = [...str2.toLowerCase()];
	const bigrams1 = charArr1.filter((e, i) => i < charArr1.length - 1)
						 .map((e, i) => e + charArr1[i + 1]);						
	const bigrams2 = charArr2.filter((e, i) => i < charArr2.length - 1)
						 .map((e, i) => e + charArr2[i + 1]);
	const intersection = new Set(bigrams1.filter(e => bigrams2.includes(e)));
	// count number of intersecting bigrams
	const intersectionCounts = [...intersection].map(bigram => Math.min(bigrams1.filter(e => e == bigram).length,
							 								            bigrams2.filter(e => e == bigram).length));
	const intersectionSize = intersectionCounts.reduce((accumulator, e) => accumulator + e, 0);
	// simply compare intersection to the shortest string only, not both
	// based on the assumption there that the longest string is most correct and complete
	let minLength = Math.min(str1.length, str2.length);
	// short names give very few bigrams, therefore need to adjust minLength in such cases- subdcract ibrams involving space to compensate for this	
	const [smallest, largest] = (str1.length > str2.length)? [str2, str1]: [str1, str2];
	if (!smallest.includes(" "))	// smallest does not include space
		{
		minLength -= largest.split(" ").length -1; // subtract false bigram counts due to likely missing space
		}	
	return intersectionSize/minLength;
	}

// new unique name routine - compare using asymmetric dice - keep variations for future lookup, keep the largest name
const hardThreshold = 0.8;
function commonNameBase(list)
	{
	let combinedStrings = {};
	// get one instance of each string
	let uniqueStrings = [...new Set(list)];
	uniqueStrings.sort((a, b) =>  b.length > a.length ? 1 : -1);	// sort strings on length, longest first
	combinedStrings = {...combinedStrings, [uniqueStrings[0]]: []};	// copy first element
	let processedNames = [uniqueStrings[0]];	// keep track of names processed
	// traverse the unique strings, longest to shortest order.
	uniqueStrings.forEach(str1 =>	 
		{
		// get a copy of already identifed strings
 		let keys = Object.keys(combinedStrings); // Simply return the keys		
		if (processedNames.includes(str1))
			{
			return;	// don't process variations already considered	
			}
		// find names that matches the most
		// first compute the distance of the name with already identified naames
		// note that it is based on the main key (longest version) - we do not check on variants
		let contenders = keys.map(str2 =>
			{
			let d1 = dice(str1, str2);
			let d2 = asymmetricDice(str1, str2);
			let score = Math.max(d1, d2);
			return {name: str2, score: score};
			});
		// identify list of items with good match
		let contendersAboveLimit = contenders.filter(({score}) => score > hardThreshold);
		// if we have no contenders, lets add this name as a new item
		if (contendersAboveLimit.length == 0)
			{
			combinedStrings = { ...combinedStrings, [str1]: []};
			processedNames.push(str1);	
			}
		else // we have a match to lets add the highest match as a variant
			{
			contendersAboveLimit.sort((a, b) => b.score - a.score);
			// pick first one
			let similarTo = contendersAboveLimit[0].name;
			combinedStrings[similarTo].push(str1);
			processedNames.push(str1);						
			}
		});
	return combinedStrings;
	}
function variationNameListToMap(combinedStrings)
	{
	let lookup = {};
	Object.keys(combinedStrings).forEach(key => 
		{
		let values = combinedStrings[key];
		lookup = {...lookup, [key]: key};
		values.forEach(variation => lookup = {...lookup, [variation]: key})
		});
	return lookup;	
	}
function mainNameList(combinedStrings)
	{
	return Object.keys(combinedStrings);	// Simply return the keys
	}
// return the number of times an item occur in the list, changes to lowercase and also remove spaces	
function frequency(item, list)
	{
	let count = 0;
	for (let s of list)
		{
		let str = (""+s).toLowerCase();		
		str = str.replace(/\s+/g, '');	
		if (str === item)
			{
			count++;
			}			
		}
	return count;
	}
// return the frequency of each untique item	
function histogram(threshold, list)
	{
	let m = new Map();		// find unique items
	let a = Array.from(new Set(list));	// reduce list to unique items
	a.map(t => (""+t).toLowerCase().split(' ').join(''))
	 .filter(str => str !== "")
	 .forEach(str => {  let f = frequency(str, list);
						if (f > threshold)
							{
							m.set(str, f);
							}
						});
	return m;
	}

// find the starttimes based on first occurrance of valid date was recorded	
function findStartTimes(uniqueDays, uniqueCodes, registrations)
	{
	// find all the registrations with valid keywords
	let validKeywordIndices = [...Array(registrations.length).keys()]	// finding indicies of entries with correct date.
				.filter(i => uniqueCodes.has(registrations[i].keyword.toLowerCase())) // must be a valid code
				.filter(i => uniqueDays.has(registrations[i].day));		// must be a valid day
	// go through list and select the first ones.
	let prevDay = ""; 
	let startTimes = validKeywordIndices.reduce((accumulator, i) => 
		{
		let day = registrations[i].day;
		if (day == prevDay)
			{
			return accumulator; // still on the same code	
			}
		prevDay = day;
		accumulator.push(registrations[i].time);
		return accumulator;
		}, []);
	return startTimes;
	}	
// search a list to see if any is sufficiently similar	
function approxInclude(item, list)
	{
	return list.some(s => (dice(s, item) > threshold));
	}
// find the indices of the most appropriate map	
function approxInindexOf(item, list)
	{
	return list.findIndex((s) => dice(s, item) > threshold);
	}	
function timeStringToSecs(t)
	{
	let ta = t.split(".");				// split into parts	
	return 3600*+ta[0]+60*ta[1]+1*ta[2]; // return the value by adding hours, minutes and seconds
	}
function timeDifference(t1, t2)
	{
	let d1 = timeStringToSecs(t1);		
	let d2 = timeStringToSecs(t2);
	return d2-d1;
	}
const capitalize = (str, lower = false) =>
  (lower ? str.toLowerCase() : str).replace(/(?:^|\s|["'([{])+\S/g, match => match.toUpperCase());

function findIndex(array,map,name)
	{
 	let index = array.indexOf(map[name]);
	if (index == -1)
		{
		console.log("serious problem with ",name," from ",map," in ",array);
		}	
	return index;
	}
// function for increasing the count associated with a key
function increaseCount(countMap,nameMap,nameVariation)
	{
	let mainName = nameMap[nameVariation];
	let currentCount = countMap[mainName];
	currentCount++;
	countMap[mainName] = currentCount;
	}

// global since it is in the interface
let thresholdLecture = 5; // was 5		
let firstTime = true;
function buildStructure(registrations)
	{
	let thresholdClassSize = thresholdLecture;
	// for the case when we have less than 5 classes - moreved unique days first as most reliable
	let uniqueDays = histogram(thresholdLecture, registrations.filter(({keyword}) => keyword !== absenceCode).map(({day}) => day));
	// check if we are in the beginning with very few elements
	// and update controls in gui
	if (firstTime)		// set to 2 first time, otherwise let html form control the valid range
		{
		let minLimitControl = document.getElementById("minForDisplay");
		minLimitControl.value = 2;	// lover the value to reflect the new max
		thresholdLecture = 2;			
		minLimitControl.max = uniqueDays.size; // making it impossible to select higher values
		firstTime = false;
		}
	// update the thresoldLecture according to probable class size
	thresholdClassSize = Math.max(...uniqueDays.values())/3;	

	let combinedNameList = commonNameBase(registrations.map(({name}) => name));			
	let nameLookupMap = variationNameListToMap(combinedNameList);
	let mainNames = mainNameList(combinedNameList);
	// to keep valid attendences
	let validAttendances = mainNames.reduce((accumulator, name) => accumulator = {...accumulator, [name]: 0} ,{});

	let uniqueNames = [...mainNames];
	let uniqueCodes = histogram(thresholdClassSize, registrations.map(({keyword}) => keyword));
	let startTimes = findStartTimes(uniqueDays, uniqueCodes, registrations);
	// remove absence keyword if used as it should not be listed in the stats
	if (uniqueCodes.has(absenceCode))
		{
		uniqueCodes.delete(absenceCode);	
		}		
	// get array version of the frequency maps
	let ud = Array.from(uniqueDays.keys());		
	let codes = Array.from(uniqueCodes.keys());			
	let rejectedReport = new Map();
	let validAbsences = mainNames.reduce((accumulator, name) => accumulator = {...accumulator, [name]: 0} ,{});
	let classAttendances = ud.reduce((accumulator,date) => accumulator = {...accumulator, [date]: 0}, {});
	let validRegistrations = [];	// a list to keep track of valid registrations to prevent recording of dublication registrations 
	let registrationDelay = {};		// keep track of registration delays for valid registrations.
	let startTime = registrations[0].time;
	let currentDate =  registrations[0].day;
	// traverse the full list find valid lectures
	registrations.forEach(({name, keyword, day, time},i) => 
		{
		if (dice(keyword, absenceCode) > 0.8)
				{
				increaseCount(validAttendances, nameLookupMap, name);		
				increaseCount(validAbsences, nameLookupMap, name);		
				return; // valid exemption
				}			
		// check that time and code are not swapped, if it is, then swap back and log warning
		if (approxInclude(name, codes) && approxInclude(keyword, uniqueNames) && dice(name, keyword) < 0.5)
			{
			// report as invalid, correct it and count it.				
			rejectedReport.set(i, "Name and code input wapped.");	
			[name, keyword] = [keyword, name];	// correct wrong input by swapping the two.
			}
		// invalid day
		if (!ud.includes(day))
			{
			rejectedReport.set(i, "Invalid day");
			return;				
			}
		// aborting cases - invalid keyword
		if (!approxInclude(keyword, codes))
			{
			rejectedReport.set(i, "Invalid keyword");					
			return;				
			}								
		// wrong code
		let idx = ud.indexOf(day);	// find the lecture index since it is a valid days
		if (dice(keyword, codes[idx]) < threshold)
			{
			rejectedReport.set(i, "Code for a different class");						
			return;
			}
		// need to also check if student registered twice
		let token = nameLookupMap[name] + day + codes[ud.indexOf(day)];
		if (validRegistrations.includes(token))
			{
			rejectedReport.set(i, "Dublicate registration");						
			return;
			}
		validRegistrations.push(token);	// add the token
		// is the registration within time?		
		if (currentDate !== day)
			{
			currentDate = day;
			startTime = time;			
			}
		let delay = timeDifference(startTime, time);
		delay /= 60; // convert from seconds to minutes	
		token = name + day + keyword;	// actual entries in registration so recall will be exact
		registrationDelay = {...registrationDelay, [token]:delay};	
		if (delay > lateRegistrationDelay)	// 30 minutes
			{
			rejectedReport.set(i,"Late registration (not rejected): " + delay + " minutes");
			}
		increaseCount(validAttendances, nameLookupMap, name);
		classAttendances[day]++;								
		});
	// finally filter the list of names according to the threshold in the interface
	mainNames = mainNames.filter(name => validAttendances[name] >= thresholdLecture || thresholdLecture == 1);
	generateReport({registrations:registrations, 
					uniqueNames:uniqueNames, 
					uniqueDays:uniqueDays, 
					startTimes:startTimes, 
					ud:ud, 
					codes:codes, 
					validAbsences:validAbsences, 
					rejectedReport:rejectedReport, 
					mainNames:mainNames, 
					validAttendances:validAttendances, 
					combinedNameList:combinedNameList, 
					nameLookupMap:nameLookupMap, 
					classAttendances:classAttendances,
					registrationDelay:registrationDelay});
	}
function generateReport({registrations, uniqueNames, uniqueDays, startTimes, ud,codes, validAbsences, rejectedReport, mainNames, validAttendances, combinedNameList, nameLookupMap, classAttendances, registrationDelay})
	{
	// clear gui
	document.getElementById("output1Id").innerHTML = "";	
	document.getElementById("navigateButtons").innerHTML = "";
	lastButtonId = "buttonstudentOverview";
	lastId = "studentOverview";
	let updatedOn = " (updated " + registrations.slice(-1).pop().day + ")";
	// attendance statistics
	let contents = [];
	let warnings = [];
	let sortedUniqueNames = [...mainNames];	// make certain we get the list of students in alphabetical order on first name
	sortedUniqueNames.sort();
	sortedUniqueNames.forEach((name, currentStudentIdx) => 
		{
		let count = validAttendances[name];
		let attendancePst = (100*count/uniqueDays.size).toFixed(1);
		let entry = {"No.":(currentStudentIdx + 1), "Name":capitalize(name), "Attendance count":count, "Percentage %":attendancePst};
		contents = [...contents, entry];
		// if student is below the limit
		const format = (parseInt(attendancePst) < parseInt(minAttendanceLimit))? "belowLimit": null;
		warnings = [...warnings, format];
		});	
	addTable("studentOverview","Attendance stats", updatedOn, contents, warnings);	
	// list of absolutely unique names in case of subtle variations.
	let allRegisteredNames = registrations.map(({name}) => name);
	let allNames = [...new Set(allRegisteredNames)];
	allNames.sort();
	let allNamesTable = allNames.map((name,i) => ({no:(i + 1), name: capitalize(name), registrations: allRegisteredNames.filter(n => n == name).length}));
	addTable("allNamesTable","All names", updatedOn, allNamesTable);

	// list of all name variations
	let nameVariationFormat = [];
	let nameVariationsTable = sortedUniqueNames.map((name, i) => 
		{
		let variations = combinedNameList[name];
		const format = (variations.length > 0)? "belowLimit": null;
		nameVariationFormat = [...nameVariationFormat, format];	
		return {id: (i + 1), name: capitalize(name), variations: capitalize(variations.join(", "))};
		});
	addTable("nameVariationsTable","Name variations", updatedOn, nameVariationsTable, nameVariationFormat);

	// ambigeous registrations
	let ambigRegTable = allNames.reduce((accumulator, name) => 
		{
		let ambigous = sortedUniqueNames.filter(comparisonName => comparisonName.includes(name));
		if (ambigous.length > 1)
			{
//			accumulator = [ ...accumulator, {name: name, ambiguities: ambigous} ];				
			accumulator = [ ...accumulator, ...ambigous.map((ambigName, i) =>  ({"Name": i> 0 ? "": capitalize(name), "Ambiguities": capitalize(ambigName)}) )];				
			}
		return accumulator;
		}, []);
	addTable("ambigousRegistration", "Ambigous registration", updatedOn, ambigRegTable);


	// make a report of ambigeous names, i.e. names that could be in several submissions.
	let nameParts = allNames.reduce((accumulator, name) => accumulator = [...accumulator, ...name.split(" ")], []);
	nameParts = [... new Set(nameParts)];	// unique items
//console.log("nameParts ", nameParts);	
	//let ambigTable = allNames.reduce((accumulator, name) => 
	let ambigTable = nameParts.reduce((accumulator, name) => 
		{
//		let ambigous = sortedUniqueNames.filter(comparisonName => comparisonName.indexOf(name) == 0);
		let ambigous = sortedUniqueNames.filter(comparisonName => comparisonName.split(" ").some(part => part.indexOf(name) == 0));
//		let ambigous = sortedUniqueNames.filter(comparisonName => comparisonName.includes(name));
//		let ambigous = sortedUniqueNames.filter(comparisonName => comparisonName.split(" ").includes(name));
		if (ambigous.length > 1)
			{
//			accumulator = [ ...accumulator, {name: name, ambiguities: ambigous} ];				
			accumulator = [ ...accumulator, ...ambigous.map((ambigName, i) =>  ({"Name": i> 0 ? "": capitalize(name), "Ambiguities": capitalize(ambigName)}) )];				
			}
		return accumulator;
		}, []);
	addTable("ambigousNames", "Ambigous names", updatedOn, ambigTable);

	// valid absences
	let absenceReport = sortedUniqueNames.map((e, i) => ({"Name":capitalize(e), "Absences":(validAbsences[e]>0)? validAbsences[e]: "", "Presences":(validAttendances[e] - validAbsences[e])}));
	addTable("absenceOverview", "Absence vs presence", updatedOn, absenceReport, warnings);
	// absence registrations -  processing indices of the names array
	// this goes back to original array - could look at the processed list, but this approach also gives the date, leave as is
	let absenceRegistrations = [...Array(registrations.length).keys()]
								 	.filter(i => (dice(registrations[i].keyword,absenceCode) > 0.8))
									.map(i => ({"Date":registrations[i].day, "Name":capitalize(registrations[i].name)}));
	addTable("absenceRegistrations", "Absence granted", updatedOn, absenceRegistrations);
	// class attendance statistics
    let classContents = [];
	let levelIndicators = [];
	// create attendance count per day
	ud.forEach((day, columnIndex) =>
		{
		let count = classAttendances[day];			
		const attPst = (100*count/uniqueNames.length).toFixed(1);
		let entry = ({"No.":columnIndex + 1, "Date": day, "Attendance":count, "%":attPst, "Keyword":codes[columnIndex], "First reg.":startTimes[columnIndex]});
		const levelFormatting = (attPst >= 90)?"highLevel":(attPst >= 70)? "mediumLevel": "belowLimit"; 
		classContents = [...classContents, entry];
		levelIndicators = [...levelIndicators, levelFormatting];
		});
	addTable("classOverview", "Class info", updatedOn, classContents, levelIndicators);
	// late registrations overview, to detect potential cheating by distributing codes
	let lateRegistrations = sortedUniqueNames.reduce((accumulator, name) => accumulator = {...accumulator, [name]:0}, {});
	// traverse registrations and count accordint to students
	registrations.forEach(({name, day, keyword}) => 
		{
		let delay = registrationDelay[name+day+keyword] ?? 0;
		if (delay > lateRegistrationDelay)	// 30 min
			{
			increaseCount(lateRegistrations, nameLookupMap, name);
			}
		});
	// crekate the table overviiew
	let freqHeading = `Frequency of late registrations (> ${lateRegistrationDelay} mins)`;
	let lateRegistrationsTable = sortedUniqueNames.reduce((accumulator, name) => 
		accumulator = lateRegistrations[name] > 0 ? [...accumulator, {"Name":capitalize(name), [freqHeading]: lateRegistrations[name]}] : accumulator, []);
	// sort on frequency, most serious first
	lateRegistrationsTable.sort((a, b) => b[freqHeading] - a[freqHeading] );
	addTable("lageRegistrations", "Late registrations", updatedOn, lateRegistrationsTable);
	
	// rejected reports
	const rejectedDetails = Array.from(rejectedReport).map(([k, v]) => ({"Registration no.":k, "Cause":v, "Time":registrations[k].datestamp, "Name":capitalize(registrations[k].name), " Keyword":registrations[k].keyword}));
	const rejectedFormat = rejectedDetails.map(({Cause}) => (Cause.includes("not rejected"))? "mediumLevel" :"belowLimit");
	addTable("rejectedOverview", "Rejected registrations", updatedOn, Array.from(rejectedDetails), rejectedFormat);
	// find the last five registrations
	addTable("last5Overview", "Last students to register", updatedOn, registrations.slice(-5));
	// all registrations
	const allFormat = registrations.map(() => null);
	Array.from(rejectedReport).forEach(([k, v]) => allFormat[k] = (v.includes("not rejected"))? "mediumLevel" :"belowLimit");
	addTable("allRegistrations", "All registrations", updatedOn, registrations.map(({datestamp, name, keyword}, i) => ({"Reg. no":(i + 1), "Timestamp":datestamp, "Name":capitalize(name), "Keyword":keyword})),allFormat);
	// clean old dropDown menu
	document.getElementById("nameInput")?.remove();	// remove existing drop down with optional chaining
	// detail student timeline - one report per student
	let datalist  = document.createElement('datalist');
	datalist.id = "nameInput";
	document.body.appendChild(datalist);

	sortedUniqueNames.forEach(currentStud => 
		{
        // insert name into the datalist
        let option  = document.createElement('option');  
        option.value = currentStud;
        datalist.appendChild(option);  

		let studentTimeline = [];
		let timelineFormat = [];
		let processedDates = [];
		let prevDay = registrations[0].day;
		registrations.forEach(({name, day, keyword, time},i) => 
			{		
			// make certain we output one line for each day without attendance//
			if (day != prevDay && uniqueDays.has(prevDay) && !processedDates.includes(prevDay))
				{
				studentTimeline = [...studentTimeline, ({"Reg. no.":"", "Date":prevDay, "Time":"", "Delay (s)":"","Name":"", "Keyword":"", "Status":"Absent from class"})];
				timelineFormat = [...timelineFormat, "belowLimit"];
				processedDates = [...processedDates, prevDay];						
				}
			prevDay = day;
			// first check if it is a relevant record
			if (nameLookupMap[name] == currentStud)
				{
				let comment ="";
				let format = "highLevel";
				let delay = "";
				// avoid only unique dates being shown to make multiple registrations per day shown
				let datePart = day;
				processedDates = [...processedDates, day];
				let keywordPart = keyword;
				if (keywordPart.match(absenceCode))
					{
					keywordPart = "";
					comment = "Valid absence registration";
					format = "mediumLevel";
					}
				if (rejectedReport.has(i))
					{
					comment = rejectedReport.get(i);
					format = "belowLimit";
					}			
				if (name+day+keyword in registrationDelay)
					{
					delay = registrationDelay[name+day+keyword];
					}
				studentTimeline = [...studentTimeline, ({"Reg. no.":i, "Date":datePart, "Time":time, "Delay (s)":delay, "Name":capitalize(name), "Keyword":keywordPart, "Status":comment})];
				timelineFormat = [...timelineFormat, format];
				}
			});
		// make certain to report last day missing if it is
		if (!processedDates.includes([...registrations].pop().day))
			{
			studentTimeline = [...studentTimeline, ({"Reg. no.":"", "Date":prevDay , "Time":"", "Delay (s)":"", "Name":"", "Keyword":"", "Status":"Absent from class"})];
			timelineFormat = [...timelineFormat, "belowLimit"];
			processedDates = [...processedDates, prevDay];						
			}
		addTable(currentStud,"Student timeline for " + currentStud, updatedOn, studentTimeline, timelineFormat, true);
		});

	// select default view
	showView("buttonstudentOverview");
	}
// generic method for building tables
function addTable(viewId, title, updatedOn, contents, formats = null, skipButton = false)
	{
	tables.set(viewId, contents);	// keep records for excel output
	titles.set(viewId, title);	
	if (contents.length == 0)
		{
		return;	// No data to display
		}
	if (!skipButton) // add a selelection button
		{		
		let button = document.createElement("button");
		button.id = "button"+viewId;
		button.innerText = title;
		button.type = "button";
		button.classList.add("disabledButton");
		button.addEventListener("click", selectView);	
		document.getElementById("navigateButtons").appendChild(button);
		}
	// create the table
	let div = document.createElement("div");
	div.id = viewId;
	div.style.display = "none";
	let h2 = document.createElement("h2");
	h2.innerText = title + updatedOn;
	div.appendChild(h2);
	let table = document.createElement("table");
	div.appendChild(table);
	// add the headers
	// Find first record with all three entries filled in in case first elements are not complete - use this to populate the header
	const noHeaders = Math.max(...contents.map(e => Object.keys(e).length));
	const headerIdx = contents.findIndex(e => Object.keys(e).length == noHeaders);
	let tr = document.createElement("tr");			
	[...Object.keys(contents[headerIdx])].forEach(e => {
						let th = document.createElement("th");
						th.innerText = e;
						tr.appendChild(th);
						});
	table.appendChild(tr);
	// add the contents
	contents.forEach((row, i) => 
		{
		let tr = document.createElement("tr");
		if (formats != null)		// add formatting
			{
			tr.classList.add(formats[i])	
			}
		[...Object.values(row)].forEach(cell => {
							let td = document.createElement("td");
							td.innerText = cell;
							tr.appendChild(td);
							});
		table.appendChild(tr);							
		});
	document.getElementById("output1Id").appendChild(div);		
	}
function buttonToggle(id, turnOn)
	{
	const e = document.getElementById(id);	
	if (turnOn)
		{
		e.classList.add("enabledButton");
		e.classList.remove("disabledButton");			
		}
	else
		{
		e.classList.remove("enabledButton");
		e.classList.add("disabledButton");	
		}
	}
function selectView(e)
	{
	let currentButtonId = e.currentTarget.id;
	showView(currentButtonId);
	}
function displayStudent()
    {
    let selected = document.getElementById("nameSelector").value;   
	// change the view
	document.getElementById(lastId).style.display = "none";	
	document.getElementById(selected).style.display = "block";	
	lastId = selected;
	}	
function showView(currentButtonId)
	{
	let id = currentButtonId.substring("button".length);
	// change visual button states
	buttonToggle("copyToClipboard", true);
	buttonToggle(lastButtonId, false);
	buttonToggle(currentButtonId, true);
	lastButtonId = currentButtonId;
	// change the view
	document.getElementById(lastId).style.display = "none";	
	document.getElementById(id).style.display = "block";	
	lastId = id;
	}
function reportToClipboard()	// copy the currently visible report to clipboard
	{
	navigator.clipboard.writeText(document.getElementById(lastId).innerText);
	buttonToggle("copyToClipboard", false);			
	}
function guiSetup()
	{
	// set up handler for file upload
	const fileSelector = document.getElementById("file-selector");
	fileSelector.addEventListener('change', (event) => loadSpreadSheet(event));
	
	absenceCode = privateLocalStoreGetItem("absenceCode");			
	minAttendanceLimit = privateLocalStoreGetItem("minAttendanceLimit");			
///	XL_row_object = JSON.parse(privateLocalStoreGetItem("sheet"));			
	if (XL_row_object == null) // if nothing in browser store
		{
		copyFromForm();	// fill variables from form
		}
	else
		{
		// get stuff from local store
		copyToForm();	
///		processSheet(XL_row_object);			
		}
	}
///function wipeData()
///	{
///	privateLocalStoreSetItem("sheet", null);			
///	}
function keepData()
	{
	privateLocalStoreSetItem("absenceCode", absenceCode);			
	privateLocalStoreSetItem("minAttendanceLimit", minAttendanceLimit);			
///	privateLocalStoreSetItem("sheet", JSON.stringify(XL_row_object));			
	}	
function copyFromForm()
	{
    absenceCode = document.getElementById("absenceCode").value;
    minAttendanceLimit = document.getElementById("minAttendanceLimit").value;
	thresholdLecture = Number(document.getElementById("minForDisplay").value);	
	}
function copyToForm()
	{
	document.getElementById("absenceCode").value = absenceCode;
    document.getElementById("minAttendanceLimit").value = minAttendanceLimit;
	}
function guiTrigger()
	{
	copyFromForm();	
	processSheet(XL_row_object);	
	}
// Starting point: event handler to ensure DOM is loaded before start is called..
window.addEventListener('DOMContentLoaded', (event) => guiSetup());