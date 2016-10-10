{-# LANGUAGE OverloadedStrings #-}

import qualified Lib            as Lib
import qualified Data.Aeson as JSON
import Control.Monad (unless)
import Data.Aeson ((.=))
import Control.Monad.Trans.Except

main :: IO ()
main = do
  Right brews <- runExceptT $
    Lib.readSheet (2 :: Int) "test/sample-inputs/Brews.xlsx" "Brews"

  let expected =
                    ("Brews" .=
                    [JSON.object ["style" .= JSON.String "ipa",
                                  "rating" .= JSON.Number 8.0,
                                  "name" .= JSON.String "hop hog",
                                  "colour" .= JSON.String "FFFFC000"
                                 ],
                      JSON.object ["style" .= JSON.String "american amber",
                                  "rating" .= JSON.Number 9.0,
                                  "name" .= JSON.String "fanta pants",
                                  "colour" .= JSON.String "FFC600AE"
                                  ]]
                    )

  unless (brews == expected) $
   do putStrLn "Bad interpretation of Brews.xlsx"
      putStrLn "Expected:"
      print expected
      putStrLn "Actual:"
      print brews
      fail "TEST FAIL"
  pure ()
