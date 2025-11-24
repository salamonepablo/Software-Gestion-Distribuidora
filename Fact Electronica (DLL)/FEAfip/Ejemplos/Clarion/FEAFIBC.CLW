  MEMBER('FEAFIP.clw')

  PRAGMA('define(init_priority=>3)')

  MAP
    MODULE('FEAFIBC0.CLW')
FEAFIBC0:DctInit             PROCEDURE
FEAFIBC0:DctKill             PROCEDURE
FEAFIBC0:FilesInit           PROCEDURE
    END
  END

DctInit PROCEDURE
  CODE
  FEAFIBC0:DctInit
  FEAFIBC0:FilesInit


DctKill PROCEDURE
  CODE
  FEAFIBC0:DctKill

