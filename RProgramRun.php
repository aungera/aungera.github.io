<html>
  <head>
    <title>PHP and R Integration Sample</title>
  </head>
  <body>
  
    <?php
      // Execute the R script within PHP code
      // Generates output as test.png image.
      
	  //$out = shell_exec("Rscript --verbose SuggestLabels.R");
	  
	  
	  //grab the date so you can concate the string with the date to access the right file
	  $date = date('Y-m-d');
	  $filename = 'All_CR_'. $date . '.pdf';
	  echo $filename;
	  echo '<br>';
	  
	  
	  $showimage = '<iframe src= \''. $filename . '\'></iframe>';
	  
	  echo '<br>';
	  
    ?>
	
	<p>yelp</p>
	
	<!-- <iframe src=$filename ></iframe> -->
    
  </body>
</html>