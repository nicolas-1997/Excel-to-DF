<h1>Excel to DF to CSV</h1>

<h2>Description</h2>

<p>It extracts the quarterly data from an Excel file and processes it to convert that data into a DF and save it in csv format.</p>


<h2>Data</h2>

<h3>The data it extracts is:</h3>

<ul>
    <li>
        <strong>Year:</strong>
        <p>
            The year of the quarter that will be used to filter the dataframe. It is obtained through the title of the excel file
        </p>
    </li>
    <li>
        <strong>Month:</strong>
        <p>The month of the quarter that will be used to filter the dataframe. It is obtained through the title of the excel file</p>
    </li>
    <li>
        <strong>Unit</strong>
        <p>Unit is a value found in the index, it is used to refer to an area within the company.</p>
    </li>
    <li>
        <strong>R</strong>
        <p>R is the unique identifier value of each of the different values of the quarter.</p>
    </li>
    <li>
        <strong>Value</strong>
        <p>They are the different values of the rows within the quarters.</p>
    </li>
    <li>
        <strong>Start of month</strong>
        <p>Beginning of the month is the first day of each month in which the registration of the values is carried out.</p>
    </li>
</ul>

<h2>How to use</h2>

<ol>
    <li>
        <p>We must define a variable with the path to the file</p>
        <code>path_file = './ReporteN2-ACHI-Real-202103.xlsx'</code>
    </li>
    <li>
        <p>The Read_file() class is instantiated passing the path to the file as a parameter.</p>
        <code>exc_file = Read_file(path_file)</code>
    </li>
    <li>
        <p>We need to use the get_list_sheet_name() method of the previously created instance, we save it in a variable.</p>
        <code>num_of_sheet = exc_file.get_list_sheet_name()</code>
    </li>
    <li>
        <p>We must define in a variable a list with the names for the columns of the dataframe.</p>
        <code>columns = ["AÑO", "MES", "UNIDAD", "R", "VALOR", "INICIO_DEL_MES"]</code>
    </li>
    <li>
        <p>In a for loop we iterate through each of the sheets in the <strong>num_of_sheet</strong> variable defined above.</p>
        <code>for sheet in range(len(num_of_sheet)):</code>
    </li>
    <li>
        <p>Inside the loop we must instantiate the class
        newDF() passing as an argument the location of the file, the current sheet we are iterating over, and the columns. We store this inside a variable called df</p>
        <code> df = newDF(path_file,num_of_sheet[sheet], columns)</code>
    </li>
    <li>
        <p>To finish we must call the save_to_csv() method to be able to save the file in csv format.</p>
        <code>df.save_to_csv()</code>
    </li>
</ol>



<h3>Note</h3>

<p>This code is useful for a specific excel file. It uses the object-oriented paradigm so that it can be scalable in the future, such as adding different saving formats.</p>
