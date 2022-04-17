First, I have Visual Studio installed, and make sure the package for VSTO add-in  is downloaded.  
To do so in the installer for Visual Studio click on modfiy

![image](https://user-images.githubusercontent.com/83418612/163732774-c590297e-0277-47d6-bce3-ded00527272c.png)

Then makesure the following blocks are selected under Desktop & Mobile

![image](https://user-images.githubusercontent.com/83418612/163732787-cb43c970-e543-417e-abdb-7b3ba2f4b7ce.png)

Then under other toolsets makesure  Office/sharepoint development is selected

![image](https://user-images.githubusercontent.com/83418612/163732853-ab8fc99f-c966-46c5-b67d-d17fdcf24aa0.png)



Open Visual Studio up and click on “Clone a Repository”.

![image](https://user-images.githubusercontent.com/83418612/163715701-f6f1a272-5874-49f7-99bd-005c830437ee.png)

Then under “Browse a Repository,” click on “GitHub” 
Then add “https://github.com/Mstroup0/SeniorDesign”  to the search bar.
Then click Clone. 
Once the Cloning is done and the code is up, In the Solution Explorer double click on the File “SeniorDesign.sln”
 
 ![image](https://user-images.githubusercontent.com/83418612/163715714-979d0630-96ab-4ac0-ad8c-d0bb9df21e80.png)

In the Solution Explorer go to the “Text” folder and expand the folder. Right click the “Dictionary.txt” file and click “Copy Full Path” and paste the file path on a notes app or anything to keep the file path for later 
Should like similar to C:\Users\....\SeniorDesign\SeniorDesign\Texts\Dictionary.txt
Since the Dictionary uses System Environment Variables do the following. Search on your computer for “Environment Variables.” Then Click “Edit System Environment Variables”

![image](https://user-images.githubusercontent.com/83418612/163715744-e42cbe67-33c3-4543-961e-e7bffd0e2f28.png)

Then click Advance and Environment Variables again. 

![image](https://user-images.githubusercontent.com/83418612/163715773-1647ff5a-3e87-4ee4-9dcf-cbee3e83315f.png)

The following Window will pop up. Click on the “New” button for the System Variables. Lower most new button 

![image](https://user-images.githubusercontent.com/83418612/163715797-82e51ba4-da8a-4ae1-8324-0db65020bea1.png)

Finally Types the following into the corresponding field:
Variable Name: PREDICTION_DICTIONARY
Variable Value: paste the file path that you got from step  here. 
Should be similar to C:\Users\....\SeniorDesign\SeniorDesign\Texts\Dictionary.txt

![image](https://user-images.githubusercontent.com/83418612/163715804-77149968-b4a6-472e-97f4-b98decb07917.png)

Click Ok, until you exit out of System Properties. 
Once done go back to Visual Studio and run the code. To do so just click “Start” at the top.  

![image](https://user-images.githubusercontent.com/83418612/163715829-caa32ba4-210f-4ac3-8d5c-34b0e472a8df.png)

Word should open up on your machine and there should be a SeniorDesign tab

![image](https://user-images.githubusercontent.com/83418612/163715860-3a43b82f-c5b7-4eed-8f4c-332f53cd7ef6.png)

![image](https://user-images.githubusercontent.com/83418612/163715870-4984d6e7-11e1-4a2a-aad0-e58df8a0d993.png)

If not click “File” then “options.” Finally Click on Add-ins. If SeniorDesign does not appear in any of the areas click on the drop down arrow next to COM add-ins and click disabled items. 

![image](https://user-images.githubusercontent.com/83418612/163715889-e43070c8-6eaf-40da-9739-a5515ec57a52.png)

 A window should pop up if SeniorDesign is in the list, select it and click on enable.  
You may notice the tab when you are using Word in general and not launching the code through Visual Studio. This does work, however testing of this has not been done. 
