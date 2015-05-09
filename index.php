<?php
include('functions.php'); 
my_constants();
?>
<!DOCTYPE html>
<html lang="en">

<head>

    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="description" content="">
    <meta name="author" content="">

    <title>PHP Excel Demo</title>

    <!-- Bootstrap Core CSS -->
    <link href="css/bootstrap.min.css" rel="stylesheet">

    <!-- Custom CSS -->
    <link href="css/logo-nav.css" rel="stylesheet">

    <!-- HTML5 Shim and Respond.js IE8 support of HTML5 elements and media queries -->
    <!-- WARNING: Respond.js doesn't work if you view the page via file:// -->
    <!--[if lt IE 9]>
        <script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
        <script src="https://oss.maxcdn.com/libs/respond.js/1.4.2/respond.min.js"></script>
    <![endif]-->

</head>

<body>

    <!-- Navigation -->
    <nav class="navbar navbar-inverse navbar-fixed-top" role="navigation">
        <div class="container">
            <!-- Brand and toggle get grouped for better mobile display -->
            <div class="navbar-header">
                <button type="button" class="navbar-toggle" data-toggle="collapse" data-target="#bs-example-navbar-collapse-1">
                    <span class="sr-only">Toggle navigation</span>
                    <span class="icon-bar"></span>
                    <span class="icon-bar"></span>
                    <span class="icon-bar"></span>
                </button>
                <a class="navbar-brand" href="#">
                    <img src="images/logo.png" alt="7codes" width="80">
                </a>
            </div>
            <!-- Collect the nav links, forms, and other content for toggling -->
            <div class="collapse navbar-collapse" id="bs-example-navbar-collapse-1">
                <ul class="nav navbar-nav pull-right">
                    <li>
                        <a href="http://www.7codes.info">Back</a>
                    </li>
                </ul>
            </div>
            <!-- /.navbar-collapse -->
        </div>
        <!-- /.container -->
    </nav>

    <!-- Page Content -->
    <div class="container">
        <div class="row">
            <div class="col-lg-12">

                <table class="table table-striped table-bordered">
                    <caption>Excel 
                        <a href="javascript:void(0)" id="xlscreation">From Ajax</a>
                        <a href="direct.php" >From Direct</a>
                    </caption>
                   <thead>
                      <tr>
                        <th>Brand Icon</th>
                        <th>Comapany</th>
                        <th>Rank</th>
                        <th>Link</th>
                      </tr>
                   </thead>
                   <tbody>
                    <?php 
                        $reportdetails = report_details(1);
                        foreach($reportdetails as $value){
                    ?>
                      <tr>
                        <td><?php echo $value['BrandIcon']; ?></td>
                        <td><?php echo $value['Comapany']; ?></td>
                        <td><?php echo $value['Rank']; ?></td>
                        <td><?php echo $value['Link']; ?></td>
                      </tr>
                    <?php } ?>
                   </tbody>
                </table>
            </div>
        </div>
    </div>
    <!-- /.container -->

    

    <!-- jQuery -->
    <script src="js/jquery.js"></script>

    <!-- Bootstrap Core JavaScript -->
    <script src="js/bootstrap.min.js"></script>

    <script type="text/javascript">
        var SITEURL = "<?php echo SITEURL; ?>";
        $(document).ready(function(){
            $( "#xlscreation" ).on( "click", function(e) {
                e.preventDefault();
                $.ajax({
                   type: "POST",
                   url: "ajax.php",
                   success: function(msg){
                        var data = JSON.parse(msg);
                        window.open(SITEURL+"reports/"+data.filename, '_blank');
                   }
               });
            });
        });
    </script>

</body>

</html>
