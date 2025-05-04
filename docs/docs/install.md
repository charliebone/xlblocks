# Installing the XlBlocks Excel Add-In

Follow these steps to install the XlBlocks Excel Add-In on your system:

---

## Step 1: Download the Add-In

1. Visit the [XlBlocks GitHub repository](https://github.com/charliebone/xlblocks).
2. Navigate to the **Releases** section.
3. Download the `.xll` file that matches your system's bitness:
    - **32-bit Excel**: Download `XlBlocks-AddIn-packed.xll`.
    - **64-bit Excel**: Download `XlBlocks-AddIn64-packed.xll`.
> **Note:** If you're unsure about your Excel version's bitness, you can check it by going to **File > Account > About Excel** in Excel.

4. Save the file to a location on your computer where you can easily find it, such as your **Documents** folder or a dedicated **XlBlocks** folder.
5. Because the XlBlocks Add-In is relatively new, Windows may "block" the file. To unblock it:
    - Right-click on the downloaded `.xll` file.
    - Select **Properties**.
    - In the Properties window, check the box that says **Unblock** (if present).
    - Click **OK** to apply the changes.

---

## Step 2: Add the Add-In Directory as a Trusted Location

In order to ensure Excel can load the add-in without any security warnings, it is best to add the directory where you saved the `.xll` file as a trusted location.
This is not a required step, but it is recommended for a smoother experience.

1. Open Excel.
2. Go to **File > Options**.
3. In the left-hand menu, select **Trust Center**.
4. Click on **Trust Center Settings**.
5. In the Trust Center, select **Trusted Locations**.
6. Add the folder where you downloaded the `.xll` file as a trusted location:
   - Click **Add new location**.
   - Browse to the folder containing the `.xll` file and confirm.

---

## Step 3: Load the Add-In

1. Go to **File > Options > Add-Ins**.
2. At the bottom of the window, select **Excel Add-ins** from the dropdown and click **Go**.
3. In the Add-Ins dialog, click **Browse** and locate the downloaded `.xll` file.
4. Select the file and click **OK** to load the add-in.

---

## Step 4: Verify Installation

1. After loading the add-in, you should see the XlBlocks functionality available in Excel. You can verify this by checking for the XlBlocks tab in the Excel ribbon.
2. If you encounter any issues, ensure the folder containing the `.xll` file is added as a trusted location in the Trust Center and the file is "unblocked" as described in Step 1.

---

You're all set! Enjoy using XlBlocks in Excel.