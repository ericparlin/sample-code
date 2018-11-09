#!/usr/bin/perl

use strict;
use warnings;
use diagnostics;
#use Data::Dump::Streamer;
use Getopt::Long;
use Text::Iconv;
use Spreadsheet::XLSX;
use DBI;
use DBD::mysql;
use SQL::Abstract;
use List::MoreUtils;

my @options = initopt();
my @lines = parsexlsx(@options);

sub initopt {
    my ($file, $dbase, $table, $user, $pass, $verbose);
    my $host = 'localhost';
    my $port = '3306';
    GetOptions (
        "file=s"                    => \$file,
        "database|db=s"             => \$dbase,
        "host|hostname|server:s"    => \$host,
        "port:s"                    => \$port,
        "table=s"                   => \$table,
        "user|username=s"           => \$user,
        "pass|pw|password=s"        => \$pass,
        "verbose"                   => \$verbose
    )
        or die("Please check input arguments\n");
    push my @initopts, ($file, $dbase, $host, $port, $table, $user, $pass);
    return @initopts;
};

sub parsexlsx {
    my $fname = shift;
    my @headers;
    my @data;
    my $type;
    my $convert = Text::Iconv -> new ("utf-8", "windows-1251");
    my $xlsx = Spreadsheet::XLSX -> new ($fname, $convert);
    foreach my $infile (@{$xlsx -> {Worksheet}}) {
        printf("File: %s\n", $infile->{Name});
        $infile -> {MaxRow} ||= $infile -> {MinRow};
        foreach my $row ($infile -> {MinRow} .. $infile -> {MaxRow}) {
            $infile -> {MaxCol} ||= $infile -> {MinCol};
            foreach my $col ($infile -> {MinCol} ..  $infile -> {MaxCol}) {
                my $cell = $infile -> {Cells} [$row] [$col];
                if ($cell) {
                    if ($cell->{Type} eq 'Numeric') {
                        $type = 'integer'
                    } elsif ($cell->{Type} eq 'Date') {
                        $type = 'date'
                    } else {
                        $type = 'varchar'
                    }

                    if ($row == 0) {
                        push @headers, ($col, $cell -> {Val})
                    } else {
                        push @data, ($row, $col, $type, $cell -> {Val})
                    }
                }
            }
        }
    }
    createtable(\@headers, \@options);
    insertdata(\@data, \@headers, \@options);
    return;
};
sub createtable {
    my @headerdata  = @{$_[0]};
    my @databaseinfo = @{$_[1]};
    my $file        = shift @databaseinfo;
    my $database    = shift @databaseinfo;
    my $host        = shift @databaseinfo;
    my $port        = shift @databaseinfo;
    my $table       = shift @databaseinfo;
    my $user        = shift @databaseinfo;
    my $pass        = shift @databaseinfo;
    my %inputheaderTable;
    my @headervalues;
    my $dsn;
    my $dbh;
    my $sep = ',';
    my $i = 0;
    my $chnk = List::MoreUtils::natatime 2, @headerdata;
    while (my @headername = $chnk->())
    {
        $inputheaderTable{'col_' . $i}      = $headername[1];
        $i++;
        push @headervalues, ("$headername[1] VARCHAR(255)", "$sep");
        #Dump(\@headervalues)->Names('HeaderValues')->Indent(2)->Out();
    }
    pop @headervalues;

    if ($host eq 'localhost') {
        $dsn = "DBI:mysql:database=$database;host=$host";
    } else {
        $dsn = "DBI:mysql:database=$database;host=$host;port=$port";
    }

    $dbh = DBI->connect($dsn, $user, $pass, { RaiseError => 1, AutoCommit => 1 });

    my $sth = $dbh->prepare("CREATE TABLE IF NOT EXISTS $table (id INTEGER NOT NULL AUTO_INCREMENT PRIMARY KEY , @headervalues)")
        or die "prepare statement failed: $dbh->errstr()";
    $sth->execute();
    $dbh->disconnect;
}
sub insertdata {
    my @inputdataList = @{$_[0]};
    my @inputdataHeaders = @{$_[1]};
    my @databaseinfo = @{$_[2]};
    my $file = shift @databaseinfo;
    my $database = shift @databaseinfo;
    my $host = shift @databaseinfo;
    my $port = shift @databaseinfo;
    my $table = shift @databaseinfo;
    my $user = shift @databaseinfo;
    my $pass = shift @databaseinfo;
    my $dsn;
    my $dbh;
    #Dump(scalar @inputdataList)->Names('inputdataList')->Indent(2)->Out();
    my @headername;
    my @dataname;
    foreach (0..$#inputdataList/32) {
        my $chnk = List::MoreUtils::natatime 4, @inputdataList;
        while (my @input = $chnk->()) {
            push @dataname, $input[3];
        }
        $chnk = List::MoreUtils::natatime 2, @inputdataHeaders;
        while (my @hinput = $chnk->()) {
            push @headername, $hinput[1];
        }
        my %finalinput;
        @finalinput{ @headername } = @dataname;
        my ($stmt, @bind) = SQL::Abstract->new->insert($table, \%finalinput);
        print $stmt, "\n";
        print join ', ', @bind, "\n";

        if ($host eq 'localhost') {
            $dsn = "DBI:mysql:database=$database;host=$host";
        } else {
            $dsn = "DBI:mysql:database=$database;host=$host;port=$port";
        }
        $dbh = DBI->connect($dsn, $user, $pass, { RaiseError => 1, AutoCommit => 1 });

        my $sth = $dbh->prepare($stmt);
        $sth->execute(@bind);
        $dbh->disconnect;
    }
}
exit 0;
