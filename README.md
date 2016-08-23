# Create Bulk Site/Sub Sites in Office 365
This PowerShell Script is used to create bulk subsites in O365 site collection

I am working on a pre-production for my next SharePoint app and I wanted to create more than 300 sub sites including the following features. I thought i will share this code to all for re-using it. Thanks!

  - Create multiple levels of sites/subsites
  - Create site with user specific site template
  - Description of each site
  - Specify the URL for each site
  - Inherit Navigation from parent site
  - Permission Inheritance from parent site
  - Copy the permission from parent site
  - Create multiple SharePoint Groups for each site
  - Assign permission for each SharePoint Group
  - Add multiple users to each SharePoint group
  
I have attached the Power Shell script along with CSV file and request you to provide your site collection url and user name in the script before using this Power Shell Script. Please feel free to pull/issue for any changes. 

Below are some of the screen shot i would like to share. 

## Site Hierarchy with multiple levels of sites and sub sites
![screenshot1](https://cloud.githubusercontent.com/assets/12201407/17590947/84b968fe-5ff9-11e6-9a01-7b64f60afd6c.png)

=========================================================================================================================

## Created SharePoint Groups with specific permissions
![screenshot2](https://cloud.githubusercontent.com/assets/12201407/17591092/2074fe3e-5ffa-11e6-8a2e-b62a1d061a12.png)

=========================================================================================================================

## Added multiple users to SharePoint Groups
![screenshot3](https://cloud.githubusercontent.com/assets/12201407/17590973/9eda372c-5ff9-11e6-94ed-0e4fe8b61472.png)

=========================================================================================================================
## Sample CSV file
![sample csv](https://cloud.githubusercontent.com/assets/12201407/17883130/57ec314c-692e-11e6-903f-595b0b10315f.png)

