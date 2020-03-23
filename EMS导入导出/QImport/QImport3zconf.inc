{ -------------------------------------------------------------------- }

{$DEFINE MAX_MATCH_IS_258}

{ Compile with -DMAXSEG_64K if the alloc function cannot allocate more
  than 64k bytes at a time (needed on systems with 16-bit int). }

{- $DEFINE MAXSEG_64K}
{$IFDEF VER70}
  {$DEFINE MAXSEG_64K}
{$ENDIF}

{$UNDEF DYNAMIC_CRC_TABLE}
{$UNDEF FASTEST}
{$define patch112}        { apply patch from the zlib home page }
{ -------------------------------------------------------------------- }
{$DEFINE Delphi32}

{$IFDEF DPMI}
  {$DEFINE MSDOS}
{$ENDIF}

{$IFDEF FPC}
 {$DEFINE Use32}
 {$UNDEF DPMI}
 {$UNDEF MSDOS}
 {$UNDEF UNALIGNED_OK}  { requires SizeOf(ush) = 2 ! }
 {$UNDEF MAXSEG_64K}
 {$UNDEF Delphi32}
{$ENDIF}
