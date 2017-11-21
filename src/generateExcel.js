
const MongoClient = require('mongodb').MongoClient;
const mongo_url = '';
const json2xls = require('json2xls');
const fs = require('fs');
const Excel = require('exceljs');

module.exports = {
	createExcel : function () {
		MongoClient.connect(mongo_url, function (err, db) {
			if(err) {
		    	console.log(err);    
		    }
		   
		     db.collection('user', function (err, collection) {
		     	if(err) {
		        	console.log(err);    
		        }
		        
		         collection.find().toArray(function(err, items) {
		            if(err) {
		            	console.log(err);    
		            }
		            
		            for(var i = 0; i < items.length; i++) {
		            	items[i] = mapInterests(items[i]);
		            	items[i] = Object.assign(items[i], items[i]['mapped_interest']);
		            	delete items[i].mapped_interest;
		            }

		            capitalizeKeys(items);
		           
		            console.log(items);

		            // Fields to be exported to excel. Specified here to ignore extra fields like Passphrase, etc
		            const options = {
		            	fields: {
		            		Name:'string', 
		            		Email:'string',
		            		Country:'string',
		            		Region:'string',
		            		Phone:'string',
		            		Skills:'string',
		            		'Web Design / Graphic Design':'string',
		            		Teaching:'string',
		            		Other:'string',
		            		'Event Planning':'string',
		            		'Orphan Database Management':'string',
		            		Marketing:'string',
		            		'Public Speaking':'string',
		            		'Grant Writing':'string',
		            		'Web Content Management':'string',
		            		'Corporate/Legal':'string',
		            	}
		            }
		            var xls = json2xls( items, options);
					fs.writeFileSync('../data.xlsx', xls, 'binary');
		                       
		        });
		        
		    });         
		});
	}
}

// Capitalize each key of the json for better readability
function capitalizeKeys(obj) {
	for(var i = 0; i<obj.length;i++) {

	    var a = obj[i];
	    for (var key in a) {
	        var temp; 
	        if (a.hasOwnProperty(key)) {
	          temp = a[key];
	          delete a[key];
	          a[key.charAt(0).toUpperCase() + key.substring(1)] = temp;
	        }
	    }
	    obj[i] = a;

	}
}

// Get interests in form of string or array b y keyword 'interests' and 'volenteerInterests' and spread it out in excel by interest having value Y or blank
function mapInterests(db_user) {
	
	if ('interests' in db_user && typeof db_user['interests'] === 'string' && db_user['interests'].length > 0) {
		var interest_arr = db_user['interests'].split(',');
	}

	else if ('volenteerInterests' in db_user && typeof db_user['volenteerInterests'] === 'string' && db_user['volenteerInterests'].length > 0) {
		var interest_arr = db_user['volenteerInterests'].split(',');
	}

	else if ( 'interests' in db_user && Array.isArray(db_user['interests'])) {
		var interest_arr = db_user['interests'];
	}

	else if ( 'volenteerInterests' in db_user && Array.isArray(db_user['volenteerInterests']) ) {
		var interest_arr = db_user['volenteerInterests'];
	}

	else {
		var interest_arr = [];
	}
	interests_values = [
						'Teaching', 
						'Event Planning', 
						'Orphan Database Management', 
						'Marketing', 
						'Public Speaking',
						'Grant Writing',
						'Web Design / Graphic Design',
						'Web Content Management',
						'Corporate/Legal',
						]
	var mapped_interest = {};
	for(var i = 0; i < interest_arr.length; i++) {
		mapped_interest[interest_arr[i]] = 'Y';
	}

	for (var i = 0; i < interests_values.length; i++) {
		if (interests_values[i] in mapped_interest) {
			continue;
		}
		else {
			mapped_interest[interests_values[i]] = '';
		}
	}

	db_user['mapped_interest'] = mapped_interest;
	return db_user;
}

