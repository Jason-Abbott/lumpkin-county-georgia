# Lumpkin County, Georgia
I was hired to integrate my [webCal](https://github.com/Jason-Abbott/webCal) widget with a county's existing event, rental and reservation system. 

- Created: **1999**
- Platform: Active Server Pages on Internet Information Server

# Project

The following have been identified as priority items to be completed for the Lumpkin County WebCal by October 29, 1999:

- Month, week, and day formats to include:
   - Event name
   - Event time
	- Event location
   - Event details  (location & details can be viewable on a drill-      down, if necessary in terms of space.)
- Filters for the calendar to be viewed by staff and public.
   - Staff 
      - Have read and write capability
      - Can do the following with the WebCal:
         - View all
         - View Recreation programs only
         - View Cultural Arts programs only
         - View After Hours programs only
         - View all events by facility, including private scheduled events
         - For any of the preceding views, staff should be able to:
            - Double click a blank date in the calendar to create a new entry by accessing the existing `NewEvent.asp` page.
            - Double click an entry in the WebCal to edit or reschedule it by accessing the existing `ManageEvents.asp` page.
            - Double click an entry to cancel it by accessing the existing `CancelEvent.asp` page.
   - Public
      - Have read-only capability
      - Can do the following with the WebCal in month, week, or day views:
         - View all
         - View Recreation programs only
         - View Cultural Arts programs only
         - View After Hours programs only
         - View events by facility
         - The public needs to be able to double click a day to access the existing `RentalSubmit.asp` page
		
Again, please make every effort to link the Web Cal to existing data structures and existing pages in the application.  

#### `NewEvent.asp`
At this point, we can just hyperlink staff directly to this page, without regard to items in the request collection.

#### `ManageEvents.asp`
The EventID needs to be passed as a session variable called session("programnum").

#### `CancelEvent.asp`
Same as `ManageEvents.asp`

#### `RentalSubmit.asp`
FacilityID and Date (ie. requested date) will need to be passed to this page, either in the request collection or as session variables.