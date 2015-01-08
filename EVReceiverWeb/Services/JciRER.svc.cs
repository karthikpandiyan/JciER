using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.EventReceivers;
using Microsoft.SharePoint.Client.UserProfiles;

namespace EVReceiverWeb.Services
{
    public class JciRER : IRemoteEventService
    {
        /// <summary>
        /// Handles events that occur before an action occurs, such as when a user adds or deletes a list item.
        /// </summary>
        /// <param name="properties">Holds information about the remote event.</param>
        /// <returns>Holds information returned from the remote event.</returns>
        public SPRemoteEventResult ProcessEvent(SPRemoteEventProperties properties)
        {
            SPRemoteEventResult result = new SPRemoteEventResult();

            using (ClientContext clientContext = TokenHelper.CreateRemoteEventReceiverClientContext(properties))
            {
                if (clientContext != null)
                {
                    clientContext.Load(clientContext.Web);

                    clientContext.ExecuteQuery();
                }
            }

            return result;
        }

        /// <summary>
        /// Handles events that occur after an action occurs, such as after a user adds an item to a list or deletes an item from a list.
        /// </summary>
        /// <param name="properties">Holds information about the remote event.</param>
        public void ProcessOneWayEvent(SPRemoteEventProperties properties)
        {
            /*
            using (ClientContext clientContext = TokenHelper.CreateRemoteEventReceiverClientContext(properties))
            {
                if (clientContext != null)
                {
                    clientContext.Load(clientContext.Web);
                    clientContext.ExecuteQuery();
                }
            }

           
            using (ClientContext clientContext =
        TokenHelper.CreateRemoteEventReceiverClientContext(properties))
            {
                if (clientContext != null)
                {
                    string firstName =
                        properties.ItemEventProperties.AfterProperties[
                            "Title"
                            ].ToString();

                    string lastName =
                        properties.ItemEventProperties.AfterProperties[
                            "Title"
                            ].ToString();

                    List lstContacts =
                        clientContext.Web.Lists.GetByTitle(
                            properties.ItemEventProperties.ListTitle
                        );

                    ListItem itemContact =
                        lstContacts.GetItemById(
                            properties.ItemEventProperties.ListItemId
                        );

                    itemContact["Title"] =
                        String.Format("{0} {1}", firstName, lastName);
                    itemContact.Update();

                    clientContext.ExecuteQuery();
                }
            }

             */
            ///


            // On Item Added event, the list item creation executes
            if (properties.EventType == SPRemoteEventType.ItemAdded)
            {
                using (ClientContext clientContext = TokenHelper.CreateRemoteEventReceiverClientContext(properties))
                {
                    if (clientContext != null)
                    {
                        //try
                        //{
                        clientContext.Load(clientContext.Web);
                        clientContext.ExecuteQuery();
                        List imageLibrary = clientContext.Web.Lists.GetByTitle("Jci");
                        ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                        ListItem oListItem = imageLibrary.GetItemById(properties.ItemEventProperties.ListItemId);
                        string _userLoginName = properties.ItemEventProperties.UserLoginName;
                        string firstName = properties.ItemEventProperties.AfterProperties["First"].ToString();

                        string lastName = properties.ItemEventProperties.AfterProperties["Last"].ToString();
                        string fullname = GetProfilePropertyFor(clientContext, _userLoginName, "LastName");
                        oListItem["fullname"] = firstName + " " + lastName +" " +fullname;
                        oListItem.Update();
                        clientContext.ExecuteQuery();


                        //
                        //using (ClientContext clientContext = TokenHelper.CreateRemoteEventReceiverClientContext(properties))
                        //{
                        ////if (clientContext != null)
                        ////{
                        ////    string firstName = properties.ItemEventProperties.AfterProperties["FirstName"].ToString();

                        ////    string lastName = properties.ItemEventProperties.AfterProperties["LastNamePhonetic"].ToString();

                        ////    List lstContacts = clientContext.Web.Lists.GetByTitle(properties.ItemEventProperties.ListTitle);

                        ////    ListItem itemContact = lstContacts.GetItemById(properties.ItemEventProperties.ListItemId);

                        ////    itemContact["FullName"] = String.Format("{0} {1}", firstName, lastName);
                        ////    itemContact.Update();

                        ////    clientContext.ExecuteQuery();
                        ////}
                        //  }
                        //
                        // }
                        //catch (Exception ex){
                        //    throw;
                        //}
                    }
                }
            }

        }

        /// <summary>
        /// Gets a user profile property Value for the specified user.
        /// </summary>
        /// <param name="ctx">An Authenticated ClientContext</param>
        /// <param name="userName">The name of the target user.</param>
        /// <param name="propertyName">The value of the property to get.</param>
        /// <returns><see cref="System.String"/>The specified profile property for the specified user. Will return an Empty String if the property is not available.</returns>
        public static string GetProfilePropertyFor(ClientContext ctx, string userName, string propertyName)
        {
            string _result = string.Empty;
            if (ctx != null)
            {
                //try
                //{
                //// PeopleManager class provides the methods for operations related to people
                PeopleManager peopleManager = new PeopleManager(ctx);
                //// GetUserProfilePropertyFor method is used to get a specific user profile property for a user
                var _profileProperty = peopleManager.GetUserProfilePropertyFor(userName, propertyName);
                ctx.ExecuteQuery();
                _result = _profileProperty.Value;
                //}
                //catch
                //{
                //    throw;
                //}
            }
            return _result;
        }
    }
}
