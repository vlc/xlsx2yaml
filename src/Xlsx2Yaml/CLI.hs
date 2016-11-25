{-# LANGUAGE DeriveGeneric #-}
{-# OPTIONS_GHC -Wwarn             #-}

module Xlsx2Yaml.CLI where

import qualified Data.List.NonEmpty  as NE
import qualified Data.Text           as T
import           Options.Applicative
import           Options.Generic     (Generic, ParseField (..),
                                      ParseFields (..), ParseRecord (..),
                                      getOnly)

data Xlsx2YamlOpts = Xlsx2YamlOpts
  { xlsx :: FilePath
  , sheet :: NE.NonEmpty (T.Text)
  , output :: FilePath
  , beginDataRow :: Maybe Int
  } deriving (Eq, Show, Generic)
instance ParseRecord Xlsx2YamlOpts

parseXlsx2YamlOpts :: Parser Xlsx2YamlOpts
parseXlsx2YamlOpts = parseRecord

-- Orphans, NonEmpty parsers for optparse-generic

instance ParseField a => ParseFields (NE.NonEmpty a) where
    parseFields h m = (NE.:|) <$> parseField h m <*> parseListOfField h m

instance ParseField a => ParseRecord (NE.NonEmpty a) where
  parseRecord = fmap getOnly parseRecord
