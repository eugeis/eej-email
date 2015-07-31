package ee.email.outlook.base;

import java.util.Date;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

import ee.email.outlook.OleAuto;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa210907(v=office.11).aspx">ContactItem</a>
 *      </p>
 *      <p>
 *      Properties | <a href="http://msdn.microsoft.com/en-us/library/aa211344(v=office.11).aspx">Account</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211346(v=office.11).aspx">Actions</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211383(v=office.11).aspx">Anniversary</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211393(v=office.11).aspx">Application</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211399(v=office.11).aspx">AssistantName</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211402(v=office.11).aspx">AssistantTelephoneNumber</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211408(v=office.11).aspx">Attachments</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211414(v=office.11).aspx">AutoResolvedWinner</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211425(v=office.11).aspx">BillingInformation</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211428(v=office.11).aspx">Birthday</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211433(v=office.11).aspx">Body</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211714(v=office.11).aspx">Business2TelephoneNumber</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211779(v=office.11).aspx">BusinessAddress</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211765(v=office.11).aspx">BusinessAddressCity</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211773(v=office.11).aspx">BusinessAddressCountry</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211775(v=office.11).aspx">BusinessAddressPostOfficeBox</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211774(v=office.11).aspx">BusinessAddressPostalCode</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211776(v=office.11).aspx">BusinessAddressState</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211778(v=office.11).aspx">BusinessAddressStreet</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211780(v=office.11).aspx">BusinessFaxNumber</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211781(v=office.11).aspx">BusinessHomePage</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211782(v=office.11).aspx">BusinessTelephoneNumber</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211784(v=office.11).aspx">CallbackTelephoneNumber</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211787(v=office.11).aspx">CarTelephoneNumber</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211788(v=office.11).aspx">Categories</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211792(v=office.11).aspx">Children</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211793(v=office.11).aspx">Class</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211798(v=office.11).aspx">Companies</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211799(v=office.11).aspx">CompanyAndFullName</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211800(v=office.11).aspx">CompanyLastFirstNoSpace</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211801(v=office.11).aspx">CompanyLastFirstSpaceOnly</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211802(v=office.11).aspx">CompanyMainTelephoneNumber</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211803(v=office.11).aspx">CompanyName</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211805(v=office.11).aspx">ComputerNetworkName</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211808(v=office.11).aspx">Conflicts</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211813(v=office.11).aspx">ConversationIndex</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211814(v=office.11).aspx">ConversationTopic</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211817(v=office.11).aspx">CreationTime</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211823(v=office.11).aspx">CustomerID</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211835(v=office.11).aspx">Department</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211845(v=office.11).aspx">DownloadState</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211850(v=office.11).aspx">Email1Address</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211849(v=office.11).aspx">Email1AddressType</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211851(v=office.11).aspx">Email1DisplayName</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211852(v=office.11).aspx">Email1EntryID</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211854(v=office.11).aspx">Email2Address</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211853(v=office.11).aspx">Email2AddressType</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211855(v=office.11).aspx">Email2DisplayName</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211856(v=office.11).aspx">Email2EntryID</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211858(v=office.11).aspx">Email3Address</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211857(v=office.11).aspx">Email3AddressType</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211859(v=office.11).aspx">Email3DisplayName</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211860(v=office.11).aspx">Email3EntryID</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211865(v=office.11).aspx">EntryID</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa212028(v=office.11).aspx">FTPSite</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211869(v=office.11).aspx">FileAs</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211976(v=office.11).aspx">FirstName</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa212024(v=office.11).aspx">FormDescription</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa212031(v=office.11).aspx">FullName</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa212030(v=office.11).aspx">FullNameAndCompany</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa212033(v=office.11).aspx">Gender</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa212040(v=office.11).aspx">GetInspector</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa212044(v=office.11).aspx">GovernmentIDNumber</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa212154(v=office.11).aspx">HasPicture</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171374(v=office.11).aspx">Hobby</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171375(v=office.11).aspx">Home2TelephoneNumber</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171403(v=office.11).aspx">HomeAddress</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171379(v=office.11).aspx">HomeAddressCity</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171383(v=office.11).aspx">HomeAddressCountry</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171391(v=office.11).aspx">HomeAddressPostOfficeBox</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171389(v=office.11).aspx">HomeAddressPostalCode</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171394(v=office.11).aspx">HomeAddressState</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171399(v=office.11).aspx">HomeAddressStreet</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171408(v=office.11).aspx">HomeFaxNumber</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171412(v=office.11).aspx">HomeTelephoneNumber</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171431(v=office.11).aspx">IMAddress</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171445(v=office.11).aspx">ISDNNumber</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171432(v=office.11).aspx">Importance</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171436(v=office.11).aspx">Initials</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171440(v=office.11).aspx">InternetFreeBusyAddress</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171442(v=office.11).aspx">IsConflict</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171454(v=office.11).aspx">ItemProperties</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171457(v=office.11).aspx">JobTitle</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171458(v=office.11).aspx">Journal</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171460(v=office.11).aspx">Language</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171461(v=office.11).aspx">LastFirstAndSuffix</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171464(v=office.11).aspx">LastFirstNoSpace</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171462(v=office.11).aspx">LastFirstNoSpaceAndSuffix</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171463(v=office.11).aspx">LastFirstNoSpaceCompany</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171466(v=office.11).aspx">LastFirstSpaceOnly</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171465(v=office.11).aspx">LastFirstSpaceOnlyCompany</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171467(v=office.11).aspx">LastModificationTime</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171469(v=office.11).aspx">LastName</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171468(v=office.11).aspx">LastNameAndFirstName</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171471(v=office.11).aspx">Links</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171481(v=office.11).aspx">MailingAddress</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171475(v=office.11).aspx">MailingAddressCity</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171476(v=office.11).aspx">MailingAddressCountry</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171478(v=office.11).aspx">MailingAddressPostOfficeBox</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171477(v=office.11).aspx">MailingAddressPostalCode</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171479(v=office.11).aspx">MailingAddressState</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171480(v=office.11).aspx">MailingAddressStreet</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171482(v=office.11).aspx">ManagerName</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171484(v=office.11).aspx">MarkForDownload</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171490(v=office.11).aspx">MessageClass</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171491(v=office.11).aspx">MiddleName</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171492(v=office.11).aspx">Mileage</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171494(v=office.11).aspx">MobileTelephoneNumber</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171694(v=office.11).aspx">NetMeetingAlias</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171706(v=office.11).aspx">NetMeetingServer</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171767(v=office.11).aspx">NickName</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171769(v=office.11).aspx">NoAging</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171774(v=office.11).aspx">OfficeLocation</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171784(v=office.11).aspx">OrganizationalIDNumber</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171798(v=office.11).aspx">OtherAddress</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171790(v=office.11).aspx">OtherAddressCity</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171791(v=office.11).aspx">OtherAddressCountry</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171794(v=office.11).aspx">OtherAddressPostOfficeBox</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171793(v=office.11).aspx">OtherAddressPostalCode</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171796(v=office.11).aspx">OtherAddressState</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171797(v=office.11).aspx">OtherAddressStreet</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171800(v=office.11).aspx">OtherFaxNumber</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171801(v=office.11).aspx">OtherTelephoneNumber</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171803(v=office.11).aspx">OutlookInternalVersion</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171805(v=office.11).aspx">OutlookVersion</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171813(v=office.11).aspx">PagerNumber</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171818(v=office.11).aspx">Parent</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171841(v=office.11).aspx">PersonalHomePage</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171850(v=office.11).aspx">PrimaryTelephoneNumber</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171855(v=office.11).aspx">Profession</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171857(v=office.11).aspx">RadioTelephoneNumber</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171885(v=office.11).aspx">ReferredBy</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171928(v=office.11).aspx">Saved</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171939(v=office.11).aspx">SelectedMailingAddress</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171987(v=office.11).aspx">Sensitivity</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172019(v=office.11).aspx">Session</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172040(v=office.11).aspx">Size</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172047(v=office.11).aspx">Spouse</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa212425(v=office.11).aspx">Subject</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220156(v=office.11).aspx">Suffix</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220727(v=office.11).aspx">TTYTDDTelephoneNumber</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220290(v=office.11).aspx">TelexNumber</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220408(v=office.11).aspx">Title</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221742(v=office.11).aspx">UnRead</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221746(v=office.11).aspx">User1</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221752(v=office.11).aspx">User2</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221758(v=office.11).aspx">User3</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221763(v=office.11).aspx">User4</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221768(v=office.11).aspx">UserCertificate</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221771(v=office.11).aspx">UserProperties</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221835(v=office.11).aspx">WebPage</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221860(v=office.11).aspx">YomiCompanyName</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221863(v=office.11).aspx">YomiFirstName</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221867(v=office.11).aspx">YomiLastName</a>
 *      </p>
 *      <p>
 *      Methods | <a href="http://msdn.microsoft.com/en-us/library/aa219412(v=office.11).aspx">AddPicture</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220077(v=office.11).aspx">Close</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220080(v=office.11).aspx">Copy</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220085(v=office.11).aspx">Delete</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220090(v=office.11).aspx">Display</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220095(v=office.11).aspx">ForwardAsVcard</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220127(v=office.11).aspx">Move</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220131(v=office.11).aspx">PrintOut</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220137(v=office.11).aspx">RemovePicture</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210281(v=office.11).aspx">Save</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210279(v=office.11).aspx">SaveAs</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210294(v=office.11).aspx">ShowCategoriesDialog</a>
 *      </p>
 *      <p>
 *      Events | <a href="http://msdn.microsoft.com/en-us/library/aa209975(v=office.11).aspx">AttachmentAdd</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa209976(v=office.11).aspx">AttachmentRead</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa209977(v=office.11).aspx">BeforeAttachmentSave</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa209978(v=office.11).aspx">BeforeCheckNames</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa209979(v=office.11).aspx">BeforeDelete</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171213(v=office.11).aspx">Close</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171218(v=office.11).aspx">CustomAction</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171227(v=office.11).aspx">CustomPropertyChange</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171259(v=office.11).aspx">Forward</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171315(v=office.11).aspx">Open</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171326(v=office.11).aspx">PropertyChange</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171358(v=office.11).aspx">Read</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171365(v=office.11).aspx">Reply</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171364(v=office.11).aspx">ReplyAll</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa219360(v=office.11).aspx">Send</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa219369(v=office.11).aspx">Write</a>
 *      </p>
 *      <p>
 *      Child Objects | <a href="http://msdn.microsoft.com/en-us/library/aa210886(v=office.11).aspx">Actions</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210901(v=office.11).aspx">Attachments</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210904(v=office.11).aspx">Conflicts</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210920(v=office.11).aspx">FormDescription</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210924(v=office.11).aspx">ItemProperties</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210939(v=office.11).aspx">Links</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211095(v=office.11).aspx">UserProperties</a>
 *      </p>*
 * @author eugeis
 */

public class ContactItem extends OleAuto {

  protected String account;

  protected Actions actions;

  protected Date anniversary;

  protected String assistantName;

  protected String assistantTelephoneNumber;

  protected Attachments attachments;

  protected Boolean autoResolvedWinner;

  protected String billingInformation;

  protected Date birthday;

  protected String body;

  protected String business2TelephoneNumber;

  protected String businessAddress;

  protected String businessAddressCity;

  protected String businessAddressCountry;

  protected String businessAddressPostOfficeBox;

  protected String businessAddressPostalCode;

  protected String businessAddressState;

  protected String businessAddressStreet;

  protected String businessFaxNumber;

  protected String businessHomePage;

  protected String businessTelephoneNumber;

  protected String callbackTelephoneNumber;

  protected String carTelephoneNumber;

  protected String categories;

  protected String children;

  protected String companies;

  protected String companyAndFullName;

  protected String companyLastFirstNoSpace;

  protected String companyLastFirstSpaceOnly;

  protected String companyMainTelephoneNumber;

  protected String companyName;

  protected String computerNetworkName;

  protected Conflicts conflicts;

  protected String conversationIndex;

  protected String conversationTopic;

  protected Date creationTime;

  protected String customerID;

  protected String department;

  protected OlDownloadStateEnum downloadState;

  protected String email1Address;

  protected String email1AddressType;

  protected String email1DisplayName;

  protected String email1EntryID;

  protected String email2Address;

  protected String email2AddressType;

  protected String email2DisplayName;

  protected String email2EntryID;

  protected String email3Address;

  protected String email3AddressType;

  protected String email3DisplayName;

  protected String email3EntryID;

  protected String entryID;

  protected String fTPSite;

  protected String fileAs;

  protected String firstName;

  protected FormDescription formDescription;

  protected String fullName;

  protected String fullNameAndCompany;

  protected OlGenderEnum gender;

  protected Variant getInspector;

  protected String governmentIDNumber;

  protected Boolean hasPicture;

  protected String hobby;

  protected String home2TelephoneNumber;

  protected String homeAddress;

  protected String homeAddressCity;

  protected String homeAddressCountry;

  protected String homeAddressPostOfficeBox;

  protected String homeAddressPostalCode;

  protected String homeAddressState;

  protected String homeAddressStreet;

  protected String homeFaxNumber;

  protected String homeTelephoneNumber;

  protected String iMAddress;

  protected String iSDNNumber;

  protected OlImportanceEnum importance;

  protected String initials;

  protected String internetFreeBusyAddress;

  protected Boolean isConflict;

  protected ItemProperties itemProperties;

  protected String jobTitle;

  protected Boolean journal;

  protected String language;

  protected String lastFirstAndSuffix;

  protected String lastFirstNoSpace;

  protected String lastFirstNoSpaceAndSuffix;

  protected String lastFirstNoSpaceCompany;

  protected Variant lastFirstSpaceOnly;

  protected String lastFirstSpaceOnlyCompany;

  protected Date lastModificationTime;

  protected String lastName;

  protected String lastNameAndFirstName;

  protected Links links;

  protected Variant mailingAddress;

  protected String mailingAddressCity;

  protected String mailingAddressCountry;

  protected String mailingAddressPostOfficeBox;

  protected String mailingAddressPostalCode;

  protected String mailingAddressState;

  protected String mailingAddressStreet;

  protected String managerName;

  protected OlRemoteStatusEnum markForDownload;

  protected String messageClass;

  protected String middleName;

  protected String mileage;

  protected String mobileTelephoneNumber;

  protected String netMeetingAlias;

  protected String netMeetingServer;

  protected String nickName;

  protected Boolean noAging;

  protected String officeLocation;

  protected String organizationalIDNumber;

  protected String otherAddress;

  protected String otherAddressCity;

  protected String otherAddressCountry;

  protected String otherAddressPostOfficeBox;

  protected String otherAddressPostalCode;

  protected String otherAddressState;

  protected String otherAddressStreet;

  protected String otherFaxNumber;

  protected String otherTelephoneNumber;

  protected Variant outlookInternalVersion;

  protected String outlookVersion;

  protected String pagerNumber;

  protected String personalHomePage;

  protected String primaryTelephoneNumber;

  protected String profession;

  protected String radioTelephoneNumber;

  protected String referredBy;

  protected Boolean saved;

  protected OlMailingAddressEnum selectedMailingAddress;

  protected OlSensitivityEnum sensitivity;

  protected Variant size;

  protected String spouse;

  protected String subject;

  protected String suffix;

  protected String tTYTDDTelephoneNumber;

  protected String telexNumber;

  protected String title;

  protected Boolean unRead;

  protected String user1;

  protected String user2;

  protected String user3;

  protected String user4;

  protected Variant userCertificate;

  protected UserProperties userProperties;

  protected String webPage;

  protected String yomiCompanyName;

  protected String yomiFirstName;

  protected String yomiLastName;

  public ContactItem(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate);
  }

  @Override
  public void init() {

    super.init();
    getAccount();
    getActions();
    getAnniversary();
    getAssistantName();
    getAssistantTelephoneNumber();
    getAttachments();
    getAutoResolvedWinner();
    getBillingInformation();
    getBirthday();
    getBody();
    getBusiness2TelephoneNumber();
    getBusinessAddress();
    getBusinessAddressCity();
    getBusinessAddressCountry();
    getBusinessAddressPostOfficeBox();
    getBusinessAddressPostalCode();
    getBusinessAddressState();
    getBusinessAddressStreet();
    getBusinessFaxNumber();
    getBusinessHomePage();
    getBusinessTelephoneNumber();
    getCallbackTelephoneNumber();
    getCarTelephoneNumber();
    getCategories();
    getChildren();
    getCompanies();
    getCompanyAndFullName();
    getCompanyLastFirstNoSpace();
    getCompanyLastFirstSpaceOnly();
    getCompanyMainTelephoneNumber();
    getCompanyName();
    getComputerNetworkName();
    getConflicts();
    getConversationIndex();
    getConversationTopic();
    getCreationTime();
    getCustomerID();
    getDepartment();
    getDownloadState();
    getEmail1Address();
    getEmail1AddressType();
    getEmail1DisplayName();
    getEmail1EntryID();
    getEmail2Address();
    getEmail2AddressType();
    getEmail2DisplayName();
    getEmail2EntryID();
    getEmail3Address();
    getEmail3AddressType();
    getEmail3DisplayName();
    getEmail3EntryID();
    getEntryID();
    getFTPSite();
    getFileAs();
    getFirstName();
    getFormDescription();
    getFullName();
    getFullNameAndCompany();
    getGender();
    getGetInspector();
    getGovernmentIDNumber();
    getHasPicture();
    getHobby();
    getHome2TelephoneNumber();
    getHomeAddress();
    getHomeAddressCity();
    getHomeAddressCountry();
    getHomeAddressPostOfficeBox();
    getHomeAddressPostalCode();
    getHomeAddressState();
    getHomeAddressStreet();
    getHomeFaxNumber();
    getHomeTelephoneNumber();
    getIMAddress();
    getISDNNumber();
    getImportance();
    getInitials();
    getInternetFreeBusyAddress();
    getIsConflict();
    getItemProperties();
    getJobTitle();
    getJournal();
    getLanguage();
    getLastFirstAndSuffix();
    getLastFirstNoSpace();
    getLastFirstNoSpaceAndSuffix();
    getLastFirstNoSpaceCompany();
    getLastFirstSpaceOnly();
    getLastFirstSpaceOnlyCompany();
    getLastModificationTime();
    getLastName();
    getLastNameAndFirstName();
    getLinks();
    getMailingAddress();
    getMailingAddressCity();
    getMailingAddressCountry();
    getMailingAddressPostOfficeBox();
    getMailingAddressPostalCode();
    getMailingAddressState();
    getMailingAddressStreet();
    getManagerName();
    getMarkForDownload();
    getMessageClass();
    getMiddleName();
    getMileage();
    getMobileTelephoneNumber();
    getNetMeetingAlias();
    getNetMeetingServer();
    getNickName();
    getNoAging();
    getOfficeLocation();
    getOrganizationalIDNumber();
    getOtherAddress();
    getOtherAddressCity();
    getOtherAddressCountry();
    getOtherAddressPostOfficeBox();
    getOtherAddressPostalCode();
    getOtherAddressState();
    getOtherAddressStreet();
    getOtherFaxNumber();
    getOtherTelephoneNumber();
    getOutlookInternalVersion();
    getOutlookVersion();
    getPagerNumber();
    getPersonalHomePage();
    getPrimaryTelephoneNumber();
    getProfession();
    getRadioTelephoneNumber();
    getReferredBy();
    getSaved();
    getSelectedMailingAddress();
    getSensitivity();
    getSize();
    getSpouse();
    getSubject();
    getSuffix();
    getTTYTDDTelephoneNumber();
    getTelexNumber();
    getTitle();
    getUnRead();
    getUser1();
    getUser2();
    getUser3();
    getUser4();
    getUserCertificate();
    getUserProperties();
    getWebPage();
    getYomiCompanyName();
    getYomiFirstName();
    getYomiLastName();
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211344(v=office.11).aspx">Account</a>
   */
  public String getAccount() {

    String propertyName = "Account";
    try {
      if (this.account == null) {
        this.account = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.account;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211346(v=office.11).aspx">Actions</a>
   */
  public Actions getActions() {

    String propertyName = "Actions";
    try {
      if (this.actions == null) {
        this.actions = new Actions(getPropertyAs(propertyName), initImmediate);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.actions;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211383(v=office.11).aspx">Anniversary</a>
   */
  public Date getAnniversary() {

    String propertyName = "Anniversary";
    try {
      if (this.anniversary == null) {
        this.anniversary = getDateValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.anniversary;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211399(v=office.11).aspx">AssistantName</a>
   */
  public String getAssistantName() {

    String propertyName = "AssistantName";
    try {
      if (this.assistantName == null) {
        this.assistantName = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.assistantName;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211402(v=office.11).aspx">AssistantTelephoneNumber</a>
   */
  public String getAssistantTelephoneNumber() {

    String propertyName = "AssistantTelephoneNumber";
    try {
      if (this.assistantTelephoneNumber == null) {
        this.assistantTelephoneNumber = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.assistantTelephoneNumber;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211408(v=office.11).aspx">Attachments</a>
   */
  public Attachments getAttachments() {

    String propertyName = "Attachments";
    try {
      if (this.attachments == null) {
        this.attachments = new Attachments(getPropertyAs(propertyName), initImmediate);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.attachments;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211414(v=office.11).aspx">AutoResolvedWinner</a>
   */
  public Boolean getAutoResolvedWinner() {

    String propertyName = "AutoResolvedWinner";
    try {
      if (this.autoResolvedWinner == null) {
        this.autoResolvedWinner = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.autoResolvedWinner;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211425(v=office.11).aspx">BillingInformation</a>
   */
  public String getBillingInformation() {

    String propertyName = "BillingInformation";
    try {
      if (this.billingInformation == null) {
        this.billingInformation = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.billingInformation;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211428(v=office.11).aspx">Birthday</a>
   */
  public Date getBirthday() {

    String propertyName = "Birthday";
    try {
      if (this.birthday == null) {
        this.birthday = getDateValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.birthday;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211433(v=office.11).aspx">Body</a>
   */
  public String getBody() {

    String propertyName = "Body";
    try {
      if (this.body == null) {
        this.body = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.body;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211714(v=office.11).aspx">Business2TelephoneNumber</a>
   */
  public String getBusiness2TelephoneNumber() {

    String propertyName = "Business2TelephoneNumber";
    try {
      if (this.business2TelephoneNumber == null) {
        this.business2TelephoneNumber = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.business2TelephoneNumber;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211779(v=office.11).aspx">BusinessAddress</a>
   */
  public String getBusinessAddress() {

    String propertyName = "BusinessAddress";
    try {
      if (this.businessAddress == null) {
        this.businessAddress = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.businessAddress;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211765(v=office.11).aspx">BusinessAddressCity</a>
   */
  public String getBusinessAddressCity() {

    String propertyName = "BusinessAddressCity";
    try {
      if (this.businessAddressCity == null) {
        this.businessAddressCity = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.businessAddressCity;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211773(v=office.11).aspx">BusinessAddressCountry</a>
   */
  public String getBusinessAddressCountry() {

    String propertyName = "BusinessAddressCountry";
    try {
      if (this.businessAddressCountry == null) {
        this.businessAddressCountry = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.businessAddressCountry;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211775(v=office.11).aspx">BusinessAddressPostOfficeBox</a>
   */
  public String getBusinessAddressPostOfficeBox() {

    String propertyName = "BusinessAddressPostOfficeBox";
    try {
      if (this.businessAddressPostOfficeBox == null) {
        this.businessAddressPostOfficeBox = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.businessAddressPostOfficeBox;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211774(v=office.11).aspx">BusinessAddressPostalCode</a>
   */
  public String getBusinessAddressPostalCode() {

    String propertyName = "BusinessAddressPostalCode";
    try {
      if (this.businessAddressPostalCode == null) {
        this.businessAddressPostalCode = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.businessAddressPostalCode;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211776(v=office.11).aspx">BusinessAddressState</a>
   */
  public String getBusinessAddressState() {

    String propertyName = "BusinessAddressState";
    try {
      if (this.businessAddressState == null) {
        this.businessAddressState = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.businessAddressState;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211778(v=office.11).aspx">BusinessAddressStreet</a>
   */
  public String getBusinessAddressStreet() {

    String propertyName = "BusinessAddressStreet";
    try {
      if (this.businessAddressStreet == null) {
        this.businessAddressStreet = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.businessAddressStreet;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211780(v=office.11).aspx">BusinessFaxNumber</a>
   */
  public String getBusinessFaxNumber() {

    String propertyName = "BusinessFaxNumber";
    try {
      if (this.businessFaxNumber == null) {
        this.businessFaxNumber = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.businessFaxNumber;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211781(v=office.11).aspx">BusinessHomePage</a>
   */
  public String getBusinessHomePage() {

    String propertyName = "BusinessHomePage";
    try {
      if (this.businessHomePage == null) {
        this.businessHomePage = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.businessHomePage;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211782(v=office.11).aspx">BusinessTelephoneNumber</a>
   */
  public String getBusinessTelephoneNumber() {

    String propertyName = "BusinessTelephoneNumber";
    try {
      if (this.businessTelephoneNumber == null) {
        this.businessTelephoneNumber = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.businessTelephoneNumber;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211784(v=office.11).aspx">CallbackTelephoneNumber</a>
   */
  public String getCallbackTelephoneNumber() {

    String propertyName = "CallbackTelephoneNumber";
    try {
      if (this.callbackTelephoneNumber == null) {
        this.callbackTelephoneNumber = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.callbackTelephoneNumber;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211787(v=office.11).aspx">CarTelephoneNumber</a>
   */
  public String getCarTelephoneNumber() {

    String propertyName = "CarTelephoneNumber";
    try {
      if (this.carTelephoneNumber == null) {
        this.carTelephoneNumber = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.carTelephoneNumber;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211788(v=office.11).aspx">Categories</a>
   */
  public String getCategories() {

    String propertyName = "Categories";
    try {
      if (this.categories == null) {
        this.categories = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.categories;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211792(v=office.11).aspx">Children</a>
   */
  public String getChildren() {

    String propertyName = "Children";
    try {
      if (this.children == null) {
        this.children = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.children;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211798(v=office.11).aspx">Companies</a>
   */
  public String getCompanies() {

    String propertyName = "Companies";
    try {
      if (this.companies == null) {
        this.companies = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.companies;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211799(v=office.11).aspx">CompanyAndFullName</a>
   */
  public String getCompanyAndFullName() {

    String propertyName = "CompanyAndFullName";
    try {
      if (this.companyAndFullName == null) {
        this.companyAndFullName = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.companyAndFullName;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211800(v=office.11).aspx">CompanyLastFirstNoSpace</a>
   */
  public String getCompanyLastFirstNoSpace() {

    String propertyName = "CompanyLastFirstNoSpace";
    try {
      if (this.companyLastFirstNoSpace == null) {
        this.companyLastFirstNoSpace = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.companyLastFirstNoSpace;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211801(v=office.11).aspx">CompanyLastFirstSpaceOnly</a>
   */
  public String getCompanyLastFirstSpaceOnly() {

    String propertyName = "CompanyLastFirstSpaceOnly";
    try {
      if (this.companyLastFirstSpaceOnly == null) {
        this.companyLastFirstSpaceOnly = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.companyLastFirstSpaceOnly;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211802(v=office.11).aspx">CompanyMainTelephoneNumber</a>
   */
  public String getCompanyMainTelephoneNumber() {

    String propertyName = "CompanyMainTelephoneNumber";
    try {
      if (this.companyMainTelephoneNumber == null) {
        this.companyMainTelephoneNumber = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.companyMainTelephoneNumber;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211803(v=office.11).aspx">CompanyName</a>
   */
  public String getCompanyName() {

    String propertyName = "CompanyName";
    try {
      if (this.companyName == null) {
        this.companyName = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.companyName;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211805(v=office.11).aspx">ComputerNetworkName</a>
   */
  public String getComputerNetworkName() {

    String propertyName = "ComputerNetworkName";
    try {
      if (this.computerNetworkName == null) {
        this.computerNetworkName = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.computerNetworkName;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211808(v=office.11).aspx">Conflicts</a>
   */
  public Conflicts getConflicts() {

    String propertyName = "Conflicts";
    try {
      if (this.conflicts == null) {
        this.conflicts = new Conflicts(getPropertyAs(propertyName), initImmediate);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.conflicts;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211813(v=office.11).aspx">ConversationIndex</a>
   */
  public String getConversationIndex() {

    String propertyName = "ConversationIndex";
    try {
      if (this.conversationIndex == null) {
        this.conversationIndex = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.conversationIndex;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211814(v=office.11).aspx">ConversationTopic</a>
   */
  public String getConversationTopic() {

    String propertyName = "ConversationTopic";
    try {
      if (this.conversationTopic == null) {
        this.conversationTopic = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.conversationTopic;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211817(v=office.11).aspx">CreationTime</a>
   */
  public Date getCreationTime() {

    String propertyName = "CreationTime";
    try {
      if (this.creationTime == null) {
        this.creationTime = getDateValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.creationTime;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211823(v=office.11).aspx">CustomerID</a>
   */
  public String getCustomerID() {

    String propertyName = "CustomerID";
    try {
      if (this.customerID == null) {
        this.customerID = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.customerID;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211835(v=office.11).aspx">Department</a>
   */
  public String getDepartment() {

    String propertyName = "Department";
    try {
      if (this.department == null) {
        this.department = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.department;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211845(v=office.11).aspx">DownloadState</a>
   */
  public OlDownloadStateEnum getDownloadState() {

    String propertyName = "DownloadState";
    try {
      if (this.downloadState == null) {
        this.downloadState = OlDownloadStateEnum.findEnum(getIntegerValue(propertyName));
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.downloadState;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211850(v=office.11).aspx">Email1Address</a>
   */
  public String getEmail1Address() {

    String propertyName = "Email1Address";
    try {
      if (this.email1Address == null) {
        this.email1Address = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.email1Address;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211849(v=office.11).aspx">Email1AddressType</a>
   */
  public String getEmail1AddressType() {

    String propertyName = "Email1AddressType";
    try {
      if (this.email1AddressType == null) {
        this.email1AddressType = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.email1AddressType;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211851(v=office.11).aspx">Email1DisplayName</a>
   */
  public String getEmail1DisplayName() {

    String propertyName = "Email1DisplayName";
    try {
      if (this.email1DisplayName == null) {
        this.email1DisplayName = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.email1DisplayName;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211852(v=office.11).aspx">Email1EntryID</a>
   */
  public String getEmail1EntryID() {

    String propertyName = "Email1EntryID";
    try {
      if (this.email1EntryID == null) {
        this.email1EntryID = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.email1EntryID;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211854(v=office.11).aspx">Email2Address</a>
   */
  public String getEmail2Address() {

    String propertyName = "Email2Address";
    try {
      if (this.email2Address == null) {
        this.email2Address = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.email2Address;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211853(v=office.11).aspx">Email2AddressType</a>
   */
  public String getEmail2AddressType() {

    String propertyName = "Email2AddressType";
    try {
      if (this.email2AddressType == null) {
        this.email2AddressType = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.email2AddressType;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211855(v=office.11).aspx">Email2DisplayName</a>
   */
  public String getEmail2DisplayName() {

    String propertyName = "Email2DisplayName";
    try {
      if (this.email2DisplayName == null) {
        this.email2DisplayName = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.email2DisplayName;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211856(v=office.11).aspx">Email2EntryID</a>
   */
  public String getEmail2EntryID() {

    String propertyName = "Email2EntryID";
    try {
      if (this.email2EntryID == null) {
        this.email2EntryID = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.email2EntryID;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211858(v=office.11).aspx">Email3Address</a>
   */
  public String getEmail3Address() {

    String propertyName = "Email3Address";
    try {
      if (this.email3Address == null) {
        this.email3Address = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.email3Address;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211857(v=office.11).aspx">Email3AddressType</a>
   */
  public String getEmail3AddressType() {

    String propertyName = "Email3AddressType";
    try {
      if (this.email3AddressType == null) {
        this.email3AddressType = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.email3AddressType;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211859(v=office.11).aspx">Email3DisplayName</a>
   */
  public String getEmail3DisplayName() {

    String propertyName = "Email3DisplayName";
    try {
      if (this.email3DisplayName == null) {
        this.email3DisplayName = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.email3DisplayName;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211860(v=office.11).aspx">Email3EntryID</a>
   */
  public String getEmail3EntryID() {

    String propertyName = "Email3EntryID";
    try {
      if (this.email3EntryID == null) {
        this.email3EntryID = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.email3EntryID;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211865(v=office.11).aspx">EntryID</a>
   */
  public String getEntryID() {

    String propertyName = "EntryID";
    try {
      if (this.entryID == null) {
        this.entryID = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.entryID;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa212028(v=office.11).aspx">FTPSite</a>
   */
  public String getFTPSite() {

    String propertyName = "FTPSite";
    try {
      if (this.fTPSite == null) {
        this.fTPSite = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.fTPSite;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211869(v=office.11).aspx">FileAs</a>
   */
  public String getFileAs() {

    String propertyName = "FileAs";
    try {
      if (this.fileAs == null) {
        this.fileAs = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.fileAs;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211976(v=office.11).aspx">FirstName</a>
   */
  public String getFirstName() {

    String propertyName = "FirstName";
    try {
      if (this.firstName == null) {
        this.firstName = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.firstName;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa212024(v=office.11).aspx">FormDescription</a>
   */
  public FormDescription getFormDescription() {

    String propertyName = "FormDescription";
    try {
      if (this.formDescription == null) {
        this.formDescription = new FormDescription(getPropertyAs(propertyName), initImmediate);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.formDescription;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa212031(v=office.11).aspx">FullName</a>
   */
  public String getFullName() {

    String propertyName = "FullName";
    try {
      if (this.fullName == null) {
        this.fullName = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.fullName;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa212030(v=office.11).aspx">FullNameAndCompany</a>
   */
  public String getFullNameAndCompany() {

    String propertyName = "FullNameAndCompany";
    try {
      if (this.fullNameAndCompany == null) {
        this.fullNameAndCompany = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.fullNameAndCompany;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa212033(v=office.11).aspx">Gender</a>
   */
  public OlGenderEnum getGender() {

    String propertyName = "Gender";
    try {
      if (this.gender == null) {
        this.gender = OlGenderEnum.findEnum(getIntegerValue(propertyName));
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.gender;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa212040(v=office.11).aspx">GetInspector</a>
   */
  public Variant getGetInspector() {

    String propertyName = "GetInspector";
    try {
      if (this.getInspector == null) {
        this.getInspector = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.getInspector;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa212044(v=office.11).aspx">GovernmentIDNumber</a>
   */
  public String getGovernmentIDNumber() {

    String propertyName = "GovernmentIDNumber";
    try {
      if (this.governmentIDNumber == null) {
        this.governmentIDNumber = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.governmentIDNumber;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa212154(v=office.11).aspx">HasPicture</a>
   */
  public Boolean getHasPicture() {

    String propertyName = "HasPicture";
    try {
      if (this.hasPicture == null) {
        this.hasPicture = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.hasPicture;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171374(v=office.11).aspx">Hobby</a>
   */
  public String getHobby() {

    String propertyName = "Hobby";
    try {
      if (this.hobby == null) {
        this.hobby = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.hobby;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171375(v=office.11).aspx">Home2TelephoneNumber</a>
   */
  public String getHome2TelephoneNumber() {

    String propertyName = "Home2TelephoneNumber";
    try {
      if (this.home2TelephoneNumber == null) {
        this.home2TelephoneNumber = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.home2TelephoneNumber;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171403(v=office.11).aspx">HomeAddress</a>
   */
  public String getHomeAddress() {

    String propertyName = "HomeAddress";
    try {
      if (this.homeAddress == null) {
        this.homeAddress = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.homeAddress;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171379(v=office.11).aspx">HomeAddressCity</a>
   */
  public String getHomeAddressCity() {

    String propertyName = "HomeAddressCity";
    try {
      if (this.homeAddressCity == null) {
        this.homeAddressCity = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.homeAddressCity;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171383(v=office.11).aspx">HomeAddressCountry</a>
   */
  public String getHomeAddressCountry() {

    String propertyName = "HomeAddressCountry";
    try {
      if (this.homeAddressCountry == null) {
        this.homeAddressCountry = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.homeAddressCountry;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171391(v=office.11).aspx">HomeAddressPostOfficeBox</a>
   */
  public String getHomeAddressPostOfficeBox() {

    String propertyName = "HomeAddressPostOfficeBox";
    try {
      if (this.homeAddressPostOfficeBox == null) {
        this.homeAddressPostOfficeBox = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.homeAddressPostOfficeBox;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171389(v=office.11).aspx">HomeAddressPostalCode</a>
   */
  public String getHomeAddressPostalCode() {

    String propertyName = "HomeAddressPostalCode";
    try {
      if (this.homeAddressPostalCode == null) {
        this.homeAddressPostalCode = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.homeAddressPostalCode;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171394(v=office.11).aspx">HomeAddressState</a>
   */
  public String getHomeAddressState() {

    String propertyName = "HomeAddressState";
    try {
      if (this.homeAddressState == null) {
        this.homeAddressState = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.homeAddressState;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171399(v=office.11).aspx">HomeAddressStreet</a>
   */
  public String getHomeAddressStreet() {

    String propertyName = "HomeAddressStreet";
    try {
      if (this.homeAddressStreet == null) {
        this.homeAddressStreet = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.homeAddressStreet;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171408(v=office.11).aspx">HomeFaxNumber</a>
   */
  public String getHomeFaxNumber() {

    String propertyName = "HomeFaxNumber";
    try {
      if (this.homeFaxNumber == null) {
        this.homeFaxNumber = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.homeFaxNumber;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171412(v=office.11).aspx">HomeTelephoneNumber</a>
   */
  public String getHomeTelephoneNumber() {

    String propertyName = "HomeTelephoneNumber";
    try {
      if (this.homeTelephoneNumber == null) {
        this.homeTelephoneNumber = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.homeTelephoneNumber;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171431(v=office.11).aspx">IMAddress</a>
   */
  public String getIMAddress() {

    String propertyName = "IMAddress";
    try {
      if (this.iMAddress == null) {
        this.iMAddress = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.iMAddress;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171445(v=office.11).aspx">ISDNNumber</a>
   */
  public String getISDNNumber() {

    String propertyName = "ISDNNumber";
    try {
      if (this.iSDNNumber == null) {
        this.iSDNNumber = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.iSDNNumber;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171432(v=office.11).aspx">Importance</a>
   */
  public OlImportanceEnum getImportance() {

    String propertyName = "Importance";
    try {
      if (this.importance == null) {
        this.importance = OlImportanceEnum.findEnum(getIntegerValue(propertyName));
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.importance;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171436(v=office.11).aspx">Initials</a>
   */
  public String getInitials() {

    String propertyName = "Initials";
    try {
      if (this.initials == null) {
        this.initials = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.initials;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171440(v=office.11).aspx">InternetFreeBusyAddress</a>
   */
  public String getInternetFreeBusyAddress() {

    String propertyName = "InternetFreeBusyAddress";
    try {
      if (this.internetFreeBusyAddress == null) {
        this.internetFreeBusyAddress = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.internetFreeBusyAddress;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171442(v=office.11).aspx">IsConflict</a>
   */
  public Boolean getIsConflict() {

    String propertyName = "IsConflict";
    try {
      if (this.isConflict == null) {
        this.isConflict = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.isConflict;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171454(v=office.11).aspx">ItemProperties</a>
   */
  public ItemProperties getItemProperties() {

    String propertyName = "ItemProperties";
    try {
      if (this.itemProperties == null) {
        this.itemProperties = new ItemProperties(getPropertyAs(propertyName), initImmediate);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.itemProperties;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171457(v=office.11).aspx">JobTitle</a>
   */
  public String getJobTitle() {

    String propertyName = "JobTitle";
    try {
      if (this.jobTitle == null) {
        this.jobTitle = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.jobTitle;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171458(v=office.11).aspx">Journal</a>
   */
  public Boolean getJournal() {

    String propertyName = "Journal";
    try {
      if (this.journal == null) {
        this.journal = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.journal;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171460(v=office.11).aspx">Language</a>
   */
  public String getLanguage() {

    String propertyName = "Language";
    try {
      if (this.language == null) {
        this.language = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.language;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171461(v=office.11).aspx">LastFirstAndSuffix</a>
   */
  public String getLastFirstAndSuffix() {

    String propertyName = "LastFirstAndSuffix";
    try {
      if (this.lastFirstAndSuffix == null) {
        this.lastFirstAndSuffix = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.lastFirstAndSuffix;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171464(v=office.11).aspx">LastFirstNoSpace</a>
   */
  public String getLastFirstNoSpace() {

    String propertyName = "LastFirstNoSpace";
    try {
      if (this.lastFirstNoSpace == null) {
        this.lastFirstNoSpace = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.lastFirstNoSpace;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171462(v=office.11).aspx">LastFirstNoSpaceAndSuffix</a>
   */
  public String getLastFirstNoSpaceAndSuffix() {

    String propertyName = "LastFirstNoSpaceAndSuffix";
    try {
      if (this.lastFirstNoSpaceAndSuffix == null) {
        this.lastFirstNoSpaceAndSuffix = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.lastFirstNoSpaceAndSuffix;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171463(v=office.11).aspx">LastFirstNoSpaceCompany</a>
   */
  public String getLastFirstNoSpaceCompany() {

    String propertyName = "LastFirstNoSpaceCompany";
    try {
      if (this.lastFirstNoSpaceCompany == null) {
        this.lastFirstNoSpaceCompany = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.lastFirstNoSpaceCompany;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171466(v=office.11).aspx">LastFirstSpaceOnly</a>
   */
  public Variant getLastFirstSpaceOnly() {

    String propertyName = "LastFirstSpaceOnly";
    try {
      if (this.lastFirstSpaceOnly == null) {
        this.lastFirstSpaceOnly = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.lastFirstSpaceOnly;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171465(v=office.11).aspx">LastFirstSpaceOnlyCompany</a>
   */
  public String getLastFirstSpaceOnlyCompany() {

    String propertyName = "LastFirstSpaceOnlyCompany";
    try {
      if (this.lastFirstSpaceOnlyCompany == null) {
        this.lastFirstSpaceOnlyCompany = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.lastFirstSpaceOnlyCompany;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171467(v=office.11).aspx">LastModificationTime</a>
   */
  public Date getLastModificationTime() {

    String propertyName = "LastModificationTime";
    try {
      if (this.lastModificationTime == null) {
        this.lastModificationTime = getDateValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.lastModificationTime;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171469(v=office.11).aspx">LastName</a>
   */
  public String getLastName() {

    String propertyName = "LastName";
    try {
      if (this.lastName == null) {
        this.lastName = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.lastName;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171468(v=office.11).aspx">LastNameAndFirstName</a>
   */
  public String getLastNameAndFirstName() {

    String propertyName = "LastNameAndFirstName";
    try {
      if (this.lastNameAndFirstName == null) {
        this.lastNameAndFirstName = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.lastNameAndFirstName;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171471(v=office.11).aspx">Links</a>
   */
  public Links getLinks() {

    String propertyName = "Links";
    try {
      if (this.links == null) {
        this.links = new Links(getPropertyAs(propertyName), initImmediate);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.links;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171481(v=office.11).aspx">MailingAddress</a>
   */
  public Variant getMailingAddress() {

    String propertyName = "MailingAddress";
    try {
      if (this.mailingAddress == null) {
        this.mailingAddress = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.mailingAddress;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171475(v=office.11).aspx">MailingAddressCity</a>
   */
  public String getMailingAddressCity() {

    String propertyName = "MailingAddressCity";
    try {
      if (this.mailingAddressCity == null) {
        this.mailingAddressCity = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.mailingAddressCity;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171476(v=office.11).aspx">MailingAddressCountry</a>
   */
  public String getMailingAddressCountry() {

    String propertyName = "MailingAddressCountry";
    try {
      if (this.mailingAddressCountry == null) {
        this.mailingAddressCountry = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.mailingAddressCountry;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171478(v=office.11).aspx">MailingAddressPostOfficeBox</a>
   */
  public String getMailingAddressPostOfficeBox() {

    String propertyName = "MailingAddressPostOfficeBox";
    try {
      if (this.mailingAddressPostOfficeBox == null) {
        this.mailingAddressPostOfficeBox = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.mailingAddressPostOfficeBox;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171477(v=office.11).aspx">MailingAddressPostalCode</a>
   */
  public String getMailingAddressPostalCode() {

    String propertyName = "MailingAddressPostalCode";
    try {
      if (this.mailingAddressPostalCode == null) {
        this.mailingAddressPostalCode = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.mailingAddressPostalCode;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171479(v=office.11).aspx">MailingAddressState</a>
   */
  public String getMailingAddressState() {

    String propertyName = "MailingAddressState";
    try {
      if (this.mailingAddressState == null) {
        this.mailingAddressState = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.mailingAddressState;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171480(v=office.11).aspx">MailingAddressStreet</a>
   */
  public String getMailingAddressStreet() {

    String propertyName = "MailingAddressStreet";
    try {
      if (this.mailingAddressStreet == null) {
        this.mailingAddressStreet = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.mailingAddressStreet;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171482(v=office.11).aspx">ManagerName</a>
   */
  public String getManagerName() {

    String propertyName = "ManagerName";
    try {
      if (this.managerName == null) {
        this.managerName = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.managerName;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171484(v=office.11).aspx">MarkForDownload</a>
   */
  public OlRemoteStatusEnum getMarkForDownload() {

    String propertyName = "MarkForDownload";
    try {
      if (this.markForDownload == null) {
        this.markForDownload = OlRemoteStatusEnum.findEnum(getIntegerValue(propertyName));
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.markForDownload;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171490(v=office.11).aspx">MessageClass</a>
   */
  public String getMessageClass() {

    String propertyName = "MessageClass";
    try {
      if (this.messageClass == null) {
        this.messageClass = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.messageClass;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171491(v=office.11).aspx">MiddleName</a>
   */
  public String getMiddleName() {

    String propertyName = "MiddleName";
    try {
      if (this.middleName == null) {
        this.middleName = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.middleName;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171492(v=office.11).aspx">Mileage</a>
   */
  public String getMileage() {

    String propertyName = "Mileage";
    try {
      if (this.mileage == null) {
        this.mileage = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.mileage;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171494(v=office.11).aspx">MobileTelephoneNumber</a>
   */
  public String getMobileTelephoneNumber() {

    String propertyName = "MobileTelephoneNumber";
    try {
      if (this.mobileTelephoneNumber == null) {
        this.mobileTelephoneNumber = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.mobileTelephoneNumber;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171694(v=office.11).aspx">NetMeetingAlias</a>
   */
  public String getNetMeetingAlias() {

    String propertyName = "NetMeetingAlias";
    try {
      if (this.netMeetingAlias == null) {
        this.netMeetingAlias = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.netMeetingAlias;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171706(v=office.11).aspx">NetMeetingServer</a>
   */
  public String getNetMeetingServer() {

    String propertyName = "NetMeetingServer";
    try {
      if (this.netMeetingServer == null) {
        this.netMeetingServer = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.netMeetingServer;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171767(v=office.11).aspx">NickName</a>
   */
  public String getNickName() {

    String propertyName = "NickName";
    try {
      if (this.nickName == null) {
        this.nickName = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.nickName;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171769(v=office.11).aspx">NoAging</a>
   */
  public Boolean getNoAging() {

    String propertyName = "NoAging";
    try {
      if (this.noAging == null) {
        this.noAging = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.noAging;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171774(v=office.11).aspx">OfficeLocation</a>
   */
  public String getOfficeLocation() {

    String propertyName = "OfficeLocation";
    try {
      if (this.officeLocation == null) {
        this.officeLocation = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.officeLocation;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171784(v=office.11).aspx">OrganizationalIDNumber</a>
   */
  public String getOrganizationalIDNumber() {

    String propertyName = "OrganizationalIDNumber";
    try {
      if (this.organizationalIDNumber == null) {
        this.organizationalIDNumber = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.organizationalIDNumber;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171798(v=office.11).aspx">OtherAddress</a>
   */
  public String getOtherAddress() {

    String propertyName = "OtherAddress";
    try {
      if (this.otherAddress == null) {
        this.otherAddress = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.otherAddress;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171790(v=office.11).aspx">OtherAddressCity</a>
   */
  public String getOtherAddressCity() {

    String propertyName = "OtherAddressCity";
    try {
      if (this.otherAddressCity == null) {
        this.otherAddressCity = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.otherAddressCity;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171791(v=office.11).aspx">OtherAddressCountry</a>
   */
  public String getOtherAddressCountry() {

    String propertyName = "OtherAddressCountry";
    try {
      if (this.otherAddressCountry == null) {
        this.otherAddressCountry = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.otherAddressCountry;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171794(v=office.11).aspx">OtherAddressPostOfficeBox</a>
   */
  public String getOtherAddressPostOfficeBox() {

    String propertyName = "OtherAddressPostOfficeBox";
    try {
      if (this.otherAddressPostOfficeBox == null) {
        this.otherAddressPostOfficeBox = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.otherAddressPostOfficeBox;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171793(v=office.11).aspx">OtherAddressPostalCode</a>
   */
  public String getOtherAddressPostalCode() {

    String propertyName = "OtherAddressPostalCode";
    try {
      if (this.otherAddressPostalCode == null) {
        this.otherAddressPostalCode = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.otherAddressPostalCode;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171796(v=office.11).aspx">OtherAddressState</a>
   */
  public String getOtherAddressState() {

    String propertyName = "OtherAddressState";
    try {
      if (this.otherAddressState == null) {
        this.otherAddressState = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.otherAddressState;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171797(v=office.11).aspx">OtherAddressStreet</a>
   */
  public String getOtherAddressStreet() {

    String propertyName = "OtherAddressStreet";
    try {
      if (this.otherAddressStreet == null) {
        this.otherAddressStreet = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.otherAddressStreet;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171800(v=office.11).aspx">OtherFaxNumber</a>
   */
  public String getOtherFaxNumber() {

    String propertyName = "OtherFaxNumber";
    try {
      if (this.otherFaxNumber == null) {
        this.otherFaxNumber = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.otherFaxNumber;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171801(v=office.11).aspx">OtherTelephoneNumber</a>
   */
  public String getOtherTelephoneNumber() {

    String propertyName = "OtherTelephoneNumber";
    try {
      if (this.otherTelephoneNumber == null) {
        this.otherTelephoneNumber = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.otherTelephoneNumber;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171803(v=office.11).aspx">OutlookInternalVersion</a>
   */
  public Variant getOutlookInternalVersion() {

    String propertyName = "OutlookInternalVersion";
    try {
      if (this.outlookInternalVersion == null) {
        this.outlookInternalVersion = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.outlookInternalVersion;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171805(v=office.11).aspx">OutlookVersion</a>
   */
  public String getOutlookVersion() {

    String propertyName = "OutlookVersion";
    try {
      if (this.outlookVersion == null) {
        this.outlookVersion = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.outlookVersion;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171813(v=office.11).aspx">PagerNumber</a>
   */
  public String getPagerNumber() {

    String propertyName = "PagerNumber";
    try {
      if (this.pagerNumber == null) {
        this.pagerNumber = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.pagerNumber;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171841(v=office.11).aspx">PersonalHomePage</a>
   */
  public String getPersonalHomePage() {

    String propertyName = "PersonalHomePage";
    try {
      if (this.personalHomePage == null) {
        this.personalHomePage = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.personalHomePage;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171850(v=office.11).aspx">PrimaryTelephoneNumber</a>
   */
  public String getPrimaryTelephoneNumber() {

    String propertyName = "PrimaryTelephoneNumber";
    try {
      if (this.primaryTelephoneNumber == null) {
        this.primaryTelephoneNumber = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.primaryTelephoneNumber;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171855(v=office.11).aspx">Profession</a>
   */
  public String getProfession() {

    String propertyName = "Profession";
    try {
      if (this.profession == null) {
        this.profession = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.profession;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171857(v=office.11).aspx">RadioTelephoneNumber</a>
   */
  public String getRadioTelephoneNumber() {

    String propertyName = "RadioTelephoneNumber";
    try {
      if (this.radioTelephoneNumber == null) {
        this.radioTelephoneNumber = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.radioTelephoneNumber;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171885(v=office.11).aspx">ReferredBy</a>
   */
  public String getReferredBy() {

    String propertyName = "ReferredBy";
    try {
      if (this.referredBy == null) {
        this.referredBy = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.referredBy;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171928(v=office.11).aspx">Saved</a>
   */
  public Boolean getSaved() {

    String propertyName = "Saved";
    try {
      if (this.saved == null) {
        this.saved = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.saved;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171939(v=office.11).aspx">SelectedMailingAddress</a>
   */
  public OlMailingAddressEnum getSelectedMailingAddress() {

    String propertyName = "SelectedMailingAddress";
    try {
      if (this.selectedMailingAddress == null) {
        this.selectedMailingAddress = OlMailingAddressEnum.findEnum(getIntegerValue(propertyName));
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.selectedMailingAddress;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171987(v=office.11).aspx">Sensitivity</a>
   */
  public OlSensitivityEnum getSensitivity() {

    String propertyName = "Sensitivity";
    try {
      if (this.sensitivity == null) {
        this.sensitivity = OlSensitivityEnum.findEnum(getIntegerValue(propertyName));
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.sensitivity;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa172040(v=office.11).aspx">Size</a>
   */
  public Variant getSize() {

    String propertyName = "Size";
    try {
      if (this.size == null) {
        this.size = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.size;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa172047(v=office.11).aspx">Spouse</a>
   */
  public String getSpouse() {

    String propertyName = "Spouse";
    try {
      if (this.spouse == null) {
        this.spouse = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.spouse;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa212425(v=office.11).aspx">Subject</a>
   */
  public String getSubject() {

    String propertyName = "Subject";
    try {
      if (this.subject == null) {
        this.subject = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.subject;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa220156(v=office.11).aspx">Suffix</a>
   */
  public String getSuffix() {

    String propertyName = "Suffix";
    try {
      if (this.suffix == null) {
        this.suffix = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.suffix;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa220727(v=office.11).aspx">TTYTDDTelephoneNumber</a>
   */
  public String getTTYTDDTelephoneNumber() {

    String propertyName = "TTYTDDTelephoneNumber";
    try {
      if (this.tTYTDDTelephoneNumber == null) {
        this.tTYTDDTelephoneNumber = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.tTYTDDTelephoneNumber;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa220290(v=office.11).aspx">TelexNumber</a>
   */
  public String getTelexNumber() {

    String propertyName = "TelexNumber";
    try {
      if (this.telexNumber == null) {
        this.telexNumber = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.telexNumber;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa220408(v=office.11).aspx">Title</a>
   */
  public String getTitle() {

    String propertyName = "Title";
    try {
      if (this.title == null) {
        this.title = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.title;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa221742(v=office.11).aspx">UnRead</a>
   */
  public Boolean getUnRead() {

    String propertyName = "UnRead";
    try {
      if (this.unRead == null) {
        this.unRead = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.unRead;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa221746(v=office.11).aspx">User1</a>
   */
  public String getUser1() {

    String propertyName = "User1";
    try {
      if (this.user1 == null) {
        this.user1 = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.user1;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa221752(v=office.11).aspx">User2</a>
   */
  public String getUser2() {

    String propertyName = "User2";
    try {
      if (this.user2 == null) {
        this.user2 = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.user2;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa221758(v=office.11).aspx">User3</a>
   */
  public String getUser3() {

    String propertyName = "User3";
    try {
      if (this.user3 == null) {
        this.user3 = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.user3;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa221763(v=office.11).aspx">User4</a>
   */
  public String getUser4() {

    String propertyName = "User4";
    try {
      if (this.user4 == null) {
        this.user4 = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.user4;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa221768(v=office.11).aspx">UserCertificate</a>
   */
  public Variant getUserCertificate() {

    String propertyName = "UserCertificate";
    try {
      if (this.userCertificate == null) {
        this.userCertificate = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.userCertificate;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa221771(v=office.11).aspx">UserProperties</a>
   */
  public UserProperties getUserProperties() {

    String propertyName = "UserProperties";
    try {
      if (this.userProperties == null) {
        this.userProperties = new UserProperties(getPropertyAs(propertyName), initImmediate);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.userProperties;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa221835(v=office.11).aspx">WebPage</a>
   */
  public String getWebPage() {

    String propertyName = "WebPage";
    try {
      if (this.webPage == null) {
        this.webPage = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.webPage;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa221860(v=office.11).aspx">YomiCompanyName</a>
   */
  public String getYomiCompanyName() {

    String propertyName = "YomiCompanyName";
    try {
      if (this.yomiCompanyName == null) {
        this.yomiCompanyName = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.yomiCompanyName;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa221863(v=office.11).aspx">YomiFirstName</a>
   */
  public String getYomiFirstName() {

    String propertyName = "YomiFirstName";
    try {
      if (this.yomiFirstName == null) {
        this.yomiFirstName = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.yomiFirstName;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa221867(v=office.11).aspx">YomiLastName</a>
   */
  public String getYomiLastName() {

    String propertyName = "YomiLastName";
    try {
      if (this.yomiLastName == null) {
        this.yomiLastName = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.yomiLastName;
  }

  @Override
  public void dispose() {

    super.dispose();
    if (this.actions != null) {
      this.actions.dispose();
    }
    if (this.attachments != null) {
      this.attachments.dispose();
    }
    if (this.conflicts != null) {
      this.conflicts.dispose();
    }
    if (this.formDescription != null) {
      this.formDescription.dispose();
    }
    if (this.itemProperties != null) {
      this.itemProperties.dispose();
    }
    if (this.links != null) {
      this.links.dispose();
    }
    if (this.userProperties != null) {
      this.userProperties.dispose();
    }
  }

}
