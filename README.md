# Stock-Anaylsis
### Project Overview
##### The purpose of this analysis was to programmatically analyze a data set made up of different variables (integers, strings, cost) in a more efficient manner. We began with understanding the value of one specific stock to decide if Stephens parents should invest in it or not by using VBA and conditional formatting. Using VBA, we were able to successfully find the total daily volume and yearly rerturn for the specific stock as well as any future stock and multiple years if needed.

### Results
##### There were 6 columns of information for each stock by date level that the total number of rows were more than 3,000. This was for both years. Rather than filtering and performing multiple VLOOKUPs, we built arrays that held the information we needed for the stocks: Ticker, Total Daily Volume and Return
<img width="284" alt="Screen Shot 2021-09-06 at 10 20 08 PM" src="https://user-images.githubusercontent.com/88811084/132279190-927f830b-350c-4ecd-8457-9064679bfdfe.png">

##### Once we had the information being able to pull in one run, we wanted the data to also calculate on its own. Using IF statements to see if the stock was the same name as the one row above, then continue to add the volume. If the stock below was not the same, then we know that is the last volume we need calculated for that stock.And the for loop was created by the count of stocks we had. We dicated the loop run 11 times (it starts at 0) for 12 tickers. Our statement is below:
<img width="507" alt="Screen Shot 2021-09-06 at 10 35 11 PM" src="https://user-images.githubusercontent.com/88811084/132280360-cff10c11-71ac-4009-97bb-238e9921346d.png">
take it one step further, we created a user input that allows the user to decide what year they want to analyze. It will prompt the user to type in the year and it will run in .0625 seconds. A time was set as soon as the user inputs the year and runs until the results are displayed. Beginning of sub and end of sub are shown below:
<img width="469" alt="Screen Shot 2021-09-06 at 10 36 17 PM" src="https://user-images.githubusercontent.com/88811084/132281047-f6cc44d0-7fcc-45f6-85db-5e59fba2dc66.png">
<img width="603" alt="Screen Shot 2021-09-06 at 10 44 25 PM" src="https://user-images.githubusercontent.com/88811084/132281114-49bc9a5e-2b45-4283-9ca8-56f9b619ddf1.png">
Refactoring this code slightly actually helped the data run quicker and looked more clean. One difference was the setting of the index to the array = to 0 so that it could start to store the volume. 
<img width="412" alt="Screen Shot 2021-09-06 at 10 48 33 PM" src="https://user-images.githubusercontent.com/88811084/132281595-43507149-a315-475a-bd5c-964650ac586c.png">
Before changing, it was shown below:
<img width="412" alt="Screen Shot 2021-09-06 at 10 51 40 PM" src="https://user-images.githubusercontent.com/88811084/132281714-454beb25-148d-4202-bfb0-dcac7d1a6f97.png">
Refactoring some of the codes resulted in a quick run time. The time used to be 0.28 seconds:  <img width="262" alt="Screen Shot 2021-09-06 at 10 53 57 PM" src="https://user-images.githubusercontent.com/88811084/132281804-3bc7177f-a166-4710-84e5-5371c24dd08c.png">. <img width="262" alt="Screen Shot 2021-09-06 at 10 54 37 PM" src="https://user-images.githubusercontent.com/88811084/132281861-18c89dcf-1f12-4a56-bfc0-586f1fb42109.png">
It runs as quick as 0.08 seconds now. Final conlusion, 2018 was not a good year for the stocks. We set a formatting macro be able to fomat based on the negative returns being red and positive returns being green. You can see from 2017 to 2018, it was not a good year on returns:
<img width="274" alt="Screen Shot 2021-09-06 at 10 59 23 PM" src="https://user-images.githubusercontent.com/88811084/132282204-e54b79e8-e462-417c-bf4a-85d4648cb3be.png">. <img width="274" alt="Screen Shot 2021-09-06 at 10 59 48 PM" src="https://user-images.githubusercontent.com/88811084/132282224-bec7683a-4a64-43bd-bd82-c101e24a3f3e.png">



#### Summary 
An advantage to refactoring data is that its cleaner to look for errors when you need to debug. There is less of a mess to look. A disadvantage of refactoring code is that once you change a code, not many will undersand what they're looking at or reading for because there is a possibility if more "magic numbers" occuring. 

