{-# LANGUAGE OverloadedStrings #-}

import qualified Lib            as Lib
import qualified Data.Aeson as JSON
import Control.Monad (unless)
import Data.Aeson ((.=))

main :: IO ()
main = do
  Just brews <-
    Lib.readSheet (10 :: Int) (2 :: Int) "test/sample-inputs/Brews.xlsx" "Brews"

  let expected = JSON.object [
                    ("Brews" .=
                    [JSON.object ["style" .= JSON.String "ipa",
                                  "rating" .= JSON.Number 8.0,
                                  "name" .= JSON.String "hop hog"],
                      JSON.object ["style" .= JSON.String "american amber",
                                  "rating" .= JSON.Number 9.0,
                                  "name" .= JSON.String "fanta pants"]]
                    )
                    ]

  unless (brews == expected) $
   do putStrLn "Bad interpretation of Brews.xlsx"
      putStrLn "Expected:"
      print expected
      putStrLn "Actual:"
      print brews
      fail "TEST FAIL"
  pure ()
