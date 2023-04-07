
clc; clear;

cd('C:\Users\Noah3\Coding')
addpath(genpath('C:\Users\Noah3\Coding'))

global BarManagerEmail

BarManagerEmail = 'username@events.com';
FileToOpen = uigetfile('.xlsx'); % Give me your weekly scheduling file.
CurrentSheet =  readcell(FileToOpen,'TextType','char'); % Open and name that CurrentSheet. Give me text as char not strings.
%% Finding Raw Data in the Schedule

BartenderMap = strncmp(CurrentSheet(),'EQ/BT', 5) | strncmp(CurrentSheet(),'BT', 2); % Find me all BT mentions in the schedule, and create a TF array which marks them 1
[~,BartendedEvents] = find(BartenderMap); % Find me the locations of all those 1's (BTs), and just give me the column numbers as 'BartenderEvents'

BartendedEvents = unique(BartendedEvents); % Remove the repeat column numbers (generated from events with multiple bartenders)

% We now have the columns of the Schedule we care about.
% We first want to cull all the relevant data from these columns, finding the data in them, and giving those variable names. 
% Then we can just fill cells of the Meeting Sheet with those variable names.

for i = BartendedEvents

    Event.Date = CurrentSheet(2,i);      % These go event-column by column, choosing the row that should have that datum.
    Event.Party = CurrentSheet(3,i);
    Event.EM_HW = CurrentSheet(15,i);
    Event.Location = CurrentSheet(5,i);
    Event.GuestCount = CurrentSheet(10,i);
    Event.Coordinator = CurrentSheet(9,i);

end



%% Raw Data Snipping and Clipping for Aesthetics of the Final Output

% Date Readability

ShatteredEventDate = split(Event.Date); % Gives a '1 x BartendedEvents x 2', that is, 'Tuesday' '12.15', as many layers deep as BartendedEvents. Digit-date is in 2 having come second in raw data.
 
for i = 1:size(ShatteredEventDate, 2) % If you use length(), the function will break when there is only 1 event, because length will pull the 2 dimension. Size(A,dim) pulls BartendedEvents' length reliably.
    DayNames = ShatteredEventDate{1,i,1}; 
    if strcmp(DayNames, 'Thursday')
    DayNames = DayNames(1:4);
    else
        DayNames = DayNames(1:3);
    end
    ShatteredEventDate{1,i,1} = DayNames;

end

% Now we have our day names clipped. Next we can rebuild Event.Date with the pieces flipped. 

for i = 1:size(ShatteredEventDate, 2)
    RearrangedDate = append(ShatteredEventDate{1,i,2}, ' ', ShatteredEventDate{1,i,1});  
    Event.Date{1,i} = RearrangedDate;
end

% Remove Redundant EM/HW label within row data, and just give me first
% names

for i = 1:length(Event.EM_HW)
    WithoutEM_HW = Event.EM_HW{1,i};
    WithoutEM_HW = WithoutEM_HW(4:end);
    FirstNameNoLabel = split(WithoutEM_HW);
    Event.EM_HW{1,i} = FirstNameNoLabel{1};

end

% Clip coordinator main two coordinator names to single letter, otherwise give two letters. 
% (This is impractical for a number of reasons, but makes it mimic, as an exercise, what I was writing out
% by hand.)

for i = 1:length(Event.Coordinator)
    CoordinatorAbbrief = Event.Coordinator{1,i} ;
    if strcmp(CoordinatorAbbrief, 'Bethany') | strcmp(CoordinatorAbbrief, 'Kristen')
        CoordinatorAbbrief = CoordinatorAbbrief(1:1);
    else
        CoordinatorAbbrief = CoordinatorAbbrief(1:2);
    end
    Event.Coordinator{1,i} = CoordinatorAbbrief;
    
end



%% Creating the Meeting Sheet & Pouring in our Polished Data




MeetingSheet = cell(12, 11);

MeetingSheet{1,1}  = 'Date';
MeetingSheet{1,2}  = 'Party';
MeetingSheet{1,3}  = 'EM/HW';
MeetingSheet{1,4}  = 'Location';
MeetingSheet{1,5}  = 'Whose alc.';
MeetingSheet{1,6}  = 'Who''s paying';
MeetingSheet{1,7}  = 'Guests';
MeetingSheet{1,8}  = 'Bars';
MeetingSheet{1,9}  = 'Notes';
MeetingSheet{1,10} = 'Pull Sheet';
MeetingSheet{1,11} = 'Coord.';
MeetingSheet{20,3} = 'Text';
MeetingSheet{20,4} = 'Transport';
MeetingSheet{20,10} = 'Pull';
MeetingSheet{20,11} = 'Email';
MeetingSheet{20,1} = 'Inventory';


for e = 1:length(BartendedEvents)

    MeetingSheet(e+1,1) = Event.Date(e);  % For the number of bartender events found, start at the empty row ("+1") and fill each necessary column, next row for the next event's data etc.
    MeetingSheet(e+1,2) = Event.Party(e);
    MeetingSheet(e+1,3) = Event.EM_HW(e);
    MeetingSheet(e+1,4) = Event.Location(e);
    MeetingSheet(e+1,7) = Event.GuestCount(e);
    MeetingSheet(e+1,11) = Event.Coordinator(e);
    
end
    



%%
%% Let's make a name, create the output sheet, and give it that name.


[~, Week , Ext] = fileparts(FileToOpen); % Get me the schedule file we originally fed in, ignore pathname, give me name as Week and the extension as Ext
Name = append('Bars for ', Week, Ext); % Make "Name" a string of the combination of that phrase and the week
NameNoExt = append('Bars for ', Week);

% Check for something already sitting there with our desired name and
% delete it to make room for the most recent output.

delete(['Bar Sheet Outputs\' Name])
delete(Name)


writecell(MeetingSheet, Name);




% Let's color code the names of EM/HW's for the zones we will assign to
% them (this helps for very busy weeks)

% ColorChoice = ['Blue' 'Red' 'Grey']
% 
% CurrentColorNumber = 1
% 
% for i = 1:length(Event.EM_HW)
%     cellnum = num2str(i)
%     cell = append('C',cellnum)
%     
%     xlswrite(Name,Event.EM_HW(i),'Sheet1',cell, 'Color', ColorChoice(CurrentColorNumber))
%     CurrentColorNumber = CurrentColorNumber + 1;
%     if CurrentColorNumber > 3
%         CurrentColorNumber = 1
%     end
% 
% end


%% Trying to Get It to Send the Sheet to the Listed Bar Manager's Work Email

% Let's verify this thing before we continue.

% Prompt the user for a response
response = questdlg('Would you like to see a preview of the Excel sheet?', 'Preview', 'Yes', 'No', 'No');

% Check the response and either display the preview or skip it
if strcmpi(response, 'Yes');
    
    % Display the preview here
        figure
        uitable('Data', MeetingSheet,'Position',[0 0 900 400])
        fig = gcf
        set(fig, 'Position', [500 100 900 600])





        response2 = questdlg('Does everything look good?', 'Preview', 'Yes', 'No', 'Yes');
        if strcmpi(response2, 'Yes')
            close all % If all looks good, close excel and move on with code
        else 
            close all
            movefile(Name, ['Rejected file -' Name]) % If not, rename it, stop code
            return
        end
else
    % Skip the preview
    
end
close all





% Now, , we can take that file and move it, then send
% it off.

movefile(Name, 'Bar Sheet Outputs')

WeekText = sprintf('Here is the bar meeting sheet for this week, generated automatically from the weekly layout. \n Good luck! \n \n- Noah S.'); % Special print function allows a wide variety of special characters to be slipped into a string, and be used to format the output.

props = java.lang.System.getProperties;
props.setProperty('mail.smtp.auth',                'true');
props.setProperty('mail.smtp.starttls.enable',     'true'); 
props.setProperty('mail.smtp.socketFactory.port',  '465');
props.setProperty('mail.smtp.socketFactory.class', 'javax.net.ssl.SSLSocketFactory');

setpref('Internet', 'E_mail',         'username@events.com'); % Let it be known what email to send that aforementioned email from.
setpref('Internet', 'SMTP_Username',  'username@events.com'); % Let it be known what email to send that aforementioned email from.
setpref('Internet', 'SMTP_Password',  'foobar');
setpref('Internet', 'SMTP_Server',    'smtp.gmail.com'); % Let it be known what server to send that mail through


response3 = questdlg('Would you like to email the sheet?', 'Email Prompt', 'Yes', 'No', 'No');
        if strcmpi(response3, 'Yes')
            sendmail(BarManagerEmail, NameNoExt, WeekText, Name); % Send email to (whom, with what subject, with what body, what file)
        else 
            return
        end