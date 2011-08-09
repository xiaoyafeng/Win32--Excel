use strict;
use Test::More;
use Test::Deep;
use Test::Harness;
use Win32::ExcelSimple qw/transpose_array cell2aa aa2cell/;

my $es = Win32::ExcelSimple->new(test.xls);

TODO: {   
	todo_skip "read function is not finished yet", 4 if 1;


is($es->write_cell('B1'),  1, "write data from cell B1");
is($es->write_row('A1'), (undef, 1,2,3), "write data from row A1");
is($es->write_col('B1'), (1,1,1,1),      "write data from col B1");
is($es->write_rect('B1:C2'),  ([1,2],[1,2]), "write data from a rectangle");
}
done_testing;

