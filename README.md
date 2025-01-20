![Dropzone-main](https://github.com/user-attachments/assets/6c182dac-beed-4bfa-a700-3e5aeb618011)
![Dropzone-key-features](https://github.com/GorgonUK/Modern-Dropzone-PCF/assets/59618079/8ad3d457-328e-4563-adb5-01f1958b1c4a)


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

   In the components search, type "Dropzone" in the search box, select the component, and then press "Add":
   
   ![Select Dropzone](https://github.com/GorgonUK/DropzonePCF/assets/59618079/5e1c1298-bd50-4d0e-8b10-4110ebc5dc44)

5. **Position the Component on the Form**

   The component will appear under "More components". You can now drag and drop it anywhere on the form:
   
   ![More components](https://github.com/GorgonUK/DropzonePCF/assets/59618079/fdea28a7-925f-46db-b698-8edcf26d9348)

   Drag the component to your desired location on the form. Note that the "table" and "view" selected during this step do not affect the component's functionality.

   ![Dropzone Component on Form](https://github.com/GorgonUK/Modern-Dropzone-PCF/assets/59618079/e9461aa3-fb3f-4170-a71c-7a96ef1ca421)

6. **Configure the component**

   Under "Enable SharePoint Documents?" type either Yes/No in the "static value" text input (If "No" then the SharePoint toggle is hidden):
   
   ![image](https://github.com/user-attachments/assets/fbcc9406-290e-4567-a677-63ede05a4bc6)

For any issues or suggestions, use the "Issues" tab at the top of the page.

## Q: Why can’t my users see files in the drop zone?  
**A:** Make sure users have `Read` access to the Notes (Annotation) table. Files uploaded via the drop zone are stored there.

---

## Q: System Administrators can see the files. Why not everyone else?  
**A:** System Administrators have access to all records by default. Other users need proper permissions for both the parent table (Requests) and Notes.

---

## Q: Do drop zone files follow the same permissions as the parent record?  
**A:** Yes, files inherit permissions from the parent record. Ensure the parent record’s access is configured correctly.

---

## Q: How can I check if the files are properly saved?  
**A:** Enable auditing on the Notes (Annotation) table to verify file creation and association.

---

## Q: Why does the drop zone show unrelated documents, and the subgrid display a throttling error?  
**A:** The drop zone is not currently designed to handle document libraries with over 5000 items. This limitation may cause unexpected behavior, such as displaying unrelated documents. A future improvement will address this issue.

---

## Q: How can I disable the drop zone for notes or SharePoint?  
**A:** You can disable drops directly in the edit form menu. Edit the component and set the boolean values for `enableNoteDrops` or `enableSharePointDrops` to `False`. By default, both options are set to `True`, but you can turn them off as needed.

---
