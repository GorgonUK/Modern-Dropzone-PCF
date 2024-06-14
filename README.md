# Dropzone PCF for Dynamics 365

Introducing the modern file upload component for Dynamics 365, the Dropzone PCF. This component is designed to seamlessly integrate with Dynamics 365 entities that have Notes and SharePoint Document Locations enabled. It offers a suite of functionalities including uploading, downloading, and deleting files, all while maintaining the aesthetic consistency with the new Dynamics 365 UI through the use of Fluent UI and custom CSS.

![2024-06-14_12-30-26](https://github.com/GorgonUK/Modern-Dropzone-PCF/assets/59618079/d144845e-a67b-4f64-bf4c-6535378dbe2f)


## Installation

The component supports SharePoint Document Locations and comes pre-configured to work with your existing Document Locations. For setup instructions, please refer to [Microsoft Learn's guide on enabling SharePoint document management for specific entities](https://learn.microsoft.com/en-us/power-platform/admin/enable-sharepoint-document-management-specific-entities).

### Steps to Install:

1. **Enable Notes**
   
   To utilize this component with notes, make sure that notes are enabled for your entities.
   
   ![Modern Dropzone Setup](https://github.com/GorgonUK/DropzonePCF/assets/59618079/652bdb3c-1e1f-45f4-95a9-e61cb6b2873e)

2. **Install the Solution**

   Install the Dropzone PCF solution package into your Dynamics 365 environment via the Solution Management area.

3. **Add Component to a Form**

   Open the desired form in edit mode and navigate to **Get more components**:
   
   ![Get more components](https://github.com/GorgonUK/DropzonePCF/assets/59618079/d737906e-29f2-4217-bb04-09a748ff3209)

4. **Search and Add Dropzone**

   In the components marketplace, type "Dropzone" in the search box, select the component, and then press "Add":
   
   ![Select Dropzone](https://github.com/GorgonUK/DropzonePCF/assets/59618079/5e1c1298-bd50-4d0e-8b10-4110ebc5dc44)

5. **Position the Component on the Form**

   The component will appear under "More components". You can now drag and drop it anywhere on the form:
   
   ![More components](https://github.com/GorgonUK/DropzonePCF/assets/59618079/fdea28a7-925f-46db-b698-8edcf26d9348)

   Drag the component to your desired location on the form. Note that the "table" and "view" selected during this step do not affect the component's functionality.

   ![Dropzone Component on Form](https://github.com/GorgonUK/Modern-Dropzone-PCF/assets/59618079/e9461aa3-fb3f-4170-a71c-7a96ef1ca421)

For any issues or suggestions, use the "Issues" tab at the top of the page.
