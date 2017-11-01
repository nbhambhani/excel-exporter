
const MongoClient = require('mongodb').MongoClient;
const mongo_url = '';
const json2xls = require('json2xls');
const fs = require('fs');

// Connect to the db
MongoClient.connect(mongo_url, function (err, db) {
   
     db.collection('user', function (err, collection) {
        
         collection.find().toArray(function(err, items) {
            if(err) throw err;    
            console.log(items[0]);  
            var xls = json2xls( items );
            console.log(items);
			fs.writeFileSync('data.xlsx', xls, 'binary');
            
            //better display ine xcel for interests array; in progress
            for(var i = 0; i < items.length; i++) {
            	items[i] = map_interests(items[i]);
            }
            

			y = 'Teaching,Orphan Database Management,Public Speaking';
			z = map_interests(y);
                
        });
        
    });         
});

function create_excel_data(items) {
	const data = {
				    sheets: [{
				        header: {
				            'name': 'Name',
				            'email': 'Email Address',
				            'note': 'Note',
				            'signupDate': 'Sign Up Date',
				            'phoneNumber': 'Contact Number',
				            'addressLine1': 'Address Line 1',
				            'addressLine2': 'Address Line 2',
				            'city': 'City',
				            'state': 'State',
				            'country': 'Country',
				            'profession': 'Profession',
				            'specialPassion': 'Special Passion',
				            'Grant Writing': 'Grant Writing',
				            'Graphic Design / Web design': 'Graphic Design / Web design',
				            'Orphan Database Management': 'Orphan Database Management',
				            'Web Content Management': 'Web Content Management',
				            'Marketing': 'Marketing',
				            'Corporate/Legal': 'Corporate/Legal',
				            'Event Planning': 'Event Planning',
				            'Teaching': 'Teaching',
				            'Public Speaking': 'Public Speaking',

				        },
				        items: items,
				        sheetName: 'New Volunteers',
				    }],
				    filepath: 'data.xlsx',
				};
	return data; 
}

function map_interests(db_user) {
	console.log(db_user);
	console.log(typeof db_user);
	var interest_arr = db_user['interests'].split(',');
	var mapped_interest = {};

	for(var i = 0; i < interest_arr.length; i++) {
		//console.log(interest_arr[i]);
		mapped_interest[interest_arr[i]] = 'Y';

	}
	//console.log(mapped_interest);
	db_user['interests'] = mapped_interest;
	return db_user;
} 

