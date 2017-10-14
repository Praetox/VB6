<?php
	$Link = @mysql_connect("sql06.freemysql.net", "anmbot", "vagooo") or die("MYSQL_CONNECTION_ERROR");
	@mysql_selectdb("botusers", $Link) or die("MYSQL_DATABASE_ERROR");
	$IP = $REMOTE_ADDR;
	if ($_GET['logout'] != "")
	{
		$SQL = "SELECT * FROM users WHERE Username='".$_GET['logout']."';";
		$Result = @mysql_query($SQL, $Link) or die("MYSQL_QUERY_ERROR");
		if (@mysql_num_rows($Result) > 0)
		{
			$ID = @mysql_result($Result, 0, "ID") or die("MYSQL_RESULT_ERROR");
			$SQL = "UPDATE users SET LastIP='".$IP."', Online='no' WHERE ID='".$ID."';";
			@mysql_query($SQL,$Link) or die("MYSQL_QUERY_ERROR");
			@mysql_close($Link);
			die("LOGGED_OUT");
		}
		else
		{
			@mysql_close($Link);
			die("LOGOUT_ERROR");
		}
	}
	if (isset($_POST['anmUser']) || isset($_POST['anmPass']))
	{
		if (!isset($_POST['anmUser']))
		{
			die("NO_USERNAME");
		}
		if (!isset($_POST['anmPass']))
		{
			die("NO_PASSWORD");
		}
		$SQL = "SELECT * FROM users WHERE Username='".$_POST['anmUser']."';";
		$Result = @mysql_query($SQL,$Link) or die("MYSQL_QUERY_ERROR");
		if (@mysql_num_rows($Result) > 0)
		{
			if ($_POST['anmPass'] == @mysql_result($Result, 0, "Password"))
			{
				$ID = @mysql_result($Result, 0, "ID") or die("MYSQL_RESULT_ERROR");
				$UN = @mysql_result($Result, 0, "Username") or die("MYSQL_RESULT_ERROR");
				$PW = @mysql_result($Result, 0, "Password") or die("MYSQL_RESULT_ERROR");
				if (@mysql_result($Result, 0, "Banned") == "no")
				{
					$SQL = "SELECT * FROM users WHERE Online='yes' AND Username='".$UN."';";
					$Result = @mysql_query($SQL,$Link) or die("MYSQL_QUERY_ERROR");
					if ((@mysql_num_rows($Result) > 0))
					{
						$CurIP = @mysql_result($Result, 0, "LastIP") or die("MYSQL_RESULT_ERROR");
						if ($CurIP != $IP)
						{
							$SQL = "UPDATE `users` SET Banned='yes' WHERE Username='".$UN."';";
							$Result = @mysql_query($SQL, $Link) or die("MYSQL_QUERY_ERROR");
							@mysql_close($Link);
							die("MULTIPLE_USERS");
						}
					}
					$SQL = "UPDATE users SET LastIP='".$IP."', Online='yes' WHERE ID='".$ID."';";
					@mysql_query($SQL,$Link) or die("MYSQL_QUERY_ERROR");
					@mysql_close($Link);
					die(md5($UN.$PW."verified_abcdefghijklmnop"));
				}
				else
				{
					@mysql_close($Link);
					die("USER_BANNED");
				}
			}
			else
			{
				@mysql_close($Link);
				die("INCORRECT_PASSWORD");
			}
		}
		else
		{
			@mysql_close($Link);
			die("INCORRECT_USERNAME");
		}
	}
	else
	{
		@mysql_close($Link);
		die("NO_USER_DATA");
	}
?>